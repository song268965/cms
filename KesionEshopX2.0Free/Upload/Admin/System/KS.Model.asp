<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<!--#include file="../../Plus/md5.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Admin_Model
KSCls.LoadKesion()
Set KSCls = Nothing


Class Admin_Model
        Private KS,KSCls,I
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
        %>
		<!--#include file="../../ks_cls/UserFunction.asp"-->
		<%

	   
		Public Sub LoadKesion()
		  'If Not KS.ReturnPowerResult(0, "model1") Then          '检查权限
			'Call KS.ReturnErr(1, "")
			'.End
		 ' End If
		 If KS.G("Action")="createtemplate" Then
			  response.cachecontrol="no-cache"
			  response.addHeader "pragma","no-cache"
			  response.expires=-1
			  response.expiresAbsolute=now-1
			  Response.CharSet="utf-8"
			  Dim KSUser,ChannelID,FieldXML,FieldNode
			  ChannelID=KS.ChkClng(KS.S("ChannelID"))
			  Set KSUser=New UserCls
			  call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
			  Call GetInputForm(true,ChannelID,FieldXML,FieldNode,"",0,KSUser,"")
			  Set KSUser=Nothing
			  Response.End()
		 End If
		  With Response
		    .Write "<!DOCTYPE html><html>"
			.Write "<title>模型基本参数设置</title>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<script src=""../../ks_inc/JQuery.js"" language=""JavaScript""></script>"
			.Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "</head>"
			
		
		if request("flag")<>"menu" then
			.Write "<body topmargin=""0"" leftmargin=""0"" >"
			.Write "<ul id='menu_top' class='menu_top_fixed'>"
			If KS.G("Action")="" Then
			.Write "<li class='parent' disabled><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>管理首页</span></li>"
			Else
			.Write "<li class='parent' onclick=""location.href='KS.Model.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon back'></i>管理首页</span></li>"
			End IF
			.Write "<li class='parent' onclick=""location.href='KS.Model.asp?action=total';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon num'></i>数据统计</span></li>"
			.Write "<li class='parent' onclick=""location.href='KS.Model.asp?action=Order';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><i class='icon set'></i>模型排序</span></li>"
			.Write "</ul><div class=""menu_top_fixed_height""></div>"
		else
		    .Write "<body style='background-color:#FFFFFF'>"
		end if

		  Select Case KS.G("Action")
		   Case "SetChannelParam"
				If Not KS.ReturnPowerResult(0, "KSMM10005") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
				     Session("FromFile")="KS.Model.asp"
		             Call SetChannelParam()
			    End If 
		
		   Case "Edit","Add"
		       If KS.G("Action")="Add" Then
		       If Not KS.ReturnPowerResult(0, "KSMM10000") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		            Call ChannelAddOrEdit()
			    End If
			  Else
		       If Not KS.ReturnPowerResult(0, "KSMM10001") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		            Call ChannelAddOrEdit()
			    End If
			  End If
		   Case "Order"
		        If Not KS.ReturnPowerResult(0, "KSMM10002") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		          Call ChannelOrder()
			    End If
		   Case "EditSave"
		        Call ChannelSave()
		   Case "Del"
		       If Not KS.ReturnPowerResult(0, "KSMM10002") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		          Call ChannelDel()
			    End If
		   Case "SetSearch"
		        Call SetSearch()
		   Case "ManageMenu"
		       If Not KS.ReturnPowerResult(0, "KSMM10002") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		          Call ManageMenu()
			    End If
		   Case "total"
		        If Not KS.ReturnPowerResult(0, "KSMM10004") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		          Call Total()
			    End If
		   Case Else
		       Call Main()
		  End Select
		  End With
		End Sub
		
		Sub ChannelOrder()
		  Dim RS: Set RS = Server.CreateObject("ADODB.RecordSet")
		  RS.Open "SELECT * FROM KS_Channel where channelid<>6 order by orderid asc,channelid", conn, 1, 1
		  if request("flag")="save" then
		     dim channelID:channelID=KS.FilterIds(KS.S("ChannelID"))
			 Dim i,IdArr:IDArr=Split(channelID,",")
			 For i=0 to Ubound(IDArr)
			   Conn.Execute("Update KS_Channel Set OrderID=" & KS.ChkClng(Request("orderid" & IDArr(i))) & " Where ChannelID=" & IDArr(i))
			 Next
			  Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
			  Session("FromFile")="System/KS.Model.asp"
			  ks.die "<script>top.$.dialog.alert('恭喜,批量保存模型排序成功!',function(){ top.location.reload();});</script>"

		  end if
		  With Response
				.Write "<form action='KS.Model.asp' name='myform' method='post'>"
				.Write "<input type='hidden' name='action' value='Order'/>"
				.Write "<input type='hidden' name='flag' value='save'/>"
		        .Write "<div class='pageCont2'><table width='100%' border='0' cellpadding='0' cellspacing='0'>"
				.Write " <tr class='sort'>"
				.Write "   <td align='center' width='60'>模型ID</td>"
				.Write "   <td align='center' width='100'>排序</td>"
				.Write "   <td align='center' width='100'>模型名称</td>"
				.Write "   <td align='center' width='100'>模型类型</td>"
				.Write "   <td align='center' width='150'>模型数据表</td>"
				.Write "   <td align='center' width='100'>项目名称</td>"
				.Write "   <td align='center' width='100'>项目单位</td>"
				.Write "   <td align='center'>模型备注</td>"
				.Write " </tr>"
				  Dim totalPut,MaxPerPage
				  MaxPerPage=50
					  
						 If RS.EOF And RS.BOF Then
						 Else
									totalPut = RS.RecordCount
									If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
											RS.Move (CurrentPage - 1) * MaxPerPage
									End If
									
									Do While Not RS.EOF
										 .Write "<tr class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"">"
										 .Write "<td class='splittd' style='text-align:center'>" & RS("ChannelID") &"</td>"
										 .Write "<td class='splittd' style='text-align:center'><input class='textbox' type='text' name='OrderID" & RS("ChannelID") & "' style='width:40px;text-align:center' value='" & KS.ChkClng(RS("OrderID")) &"'><input type='hidden' name='ChannelID' value='" & RS("ChannelID") & "'></td>"
										 .Write "  <td class='splittd' nowrap>" & RS("ChannelName") & "</td>"
										 .Write "   <td align='center' class='splittd'>"
										 If RS("ChannelID")<20 Then
										 .Write "<font color=#999999>系统"
										 Else
										  .Write "<font color=blue>自定义"
										 End If
										  .Write "</font></td>"
										  .Write "<td class='splittd' nowrap>" & RS("ChannelTable") & "</td>"
										  .Write "<td class='splittd' align=""center"">" & RS("ItemName") & "</td>"
										  .Write "<td class='splittd' align=""center"">" & RS("ItemUnit") & "</td>"
										  .Write "<td class='splittd' nowrap>" & RS("descript") & "</td>"

										 .Write " </tr>"
								I = I + 1
								If I >= MaxPerPage Then Exit Do
							   RS.MoveNext
							   Loop
								RS.Close
									
									
					End If
				 .Write "<tr><td colspan='2' height=""50"" class='operatingBox'><input type='submit' class='button' value='批量保存排序'> </td></form>"
				 .Write "  <td colspan='10' align='right'>"
				 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
				.Write "    </td>"
				.Write " </tr>"
				.Write "</table></div>"
				.Write "<br/><br/><br/></div>"
		 End With
		
		End Sub
		
		Sub SetSearch()
		 Dim AllowField:AllowField="'title','author','origin','keywords','intro','area'"
		 Dim ChannelID:channelid=KS.ChkClng(Request("channelid"))
		 Dim RS,FieldXML,XMLStr,Node,TemplateFile,isrewrite,maxperpage
		 Dim tj,check,xsz,ssz,title
		 If ChannelID=0 Then KS.Die "error!"
		 If Request("flag")="dosave" Then
		      dim ctid:ctid=KS.ChkClng(Request("Ctid"))
			  XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
			  XMLStr=XMLStr&"<field>" &vbcrlf
			  XMLStr=XMLStr&" <template><![CDATA[" & Request("TemplateFile") &"]]></template>" &vbcrlf
			  XMLStr=XMLStr&" <isrewrite>" & KS.ChkClng(Request("isrewrite")) &"</isrewrite>" &vbcrlf
			  XMLStr=XMLStr&" <maxperpage>" & KS.ChkClng(Request("maxperpage")) &"</maxperpage>" &vbcrlf
			  if ctid=1 then
			    XMLStr=XMLStr&" <item name=""tid"" enabled=""true"">" &vbcrlf
			  Else
			    XMLStr=XMLStr&" <item name=""tid"" enabled=""false"">" &vbcrlf
			  End If
			    XMLStr=XMLStr&"  <title>" & request("titletid") & "</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>tid</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>dy</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>0</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>0</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
			  '商城模块
			 If KS.ChkClng(KS.C_S(ChannelID,6))=5 Then
				  If KS.ChkClng(Request("cprice_member"))=1 Then
					 XMLStr=XMLStr&" <item name=""price_member"" enabled=""true"">" &vbcrlf
				  Else
					 XMLStr=XMLStr&" <item name=""price_member"" enabled=""false"">" &vbcrlf
				  End If
			    XMLStr=XMLStr&"  <title>" & request("titleprice_member") & "</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>price_member</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>fw</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>" & request("xszprice_member") &"</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>" & request("sszprice_member") &"</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
				  If KS.ChkClng(Request("cbrandid"))=1 Then
					 XMLStr=XMLStr&" <item name=""brandid"" enabled=""true"">" &vbcrlf
				  Else
					 XMLStr=XMLStr&" <item name=""brandid"" enabled=""false"">" &vbcrlf
				  End If
			    XMLStr=XMLStr&"  <title>" & request("titlebrandid") & "</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>brandid</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>dys</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>0</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>0</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
			 ElseIf KS.ChkClng(KS.C_S(ChannelID,6))=8 Then
				  If KS.ChkClng(Request("ctypeid"))=1 Then
					 XMLStr=XMLStr&" <item name=""typeid"" enabled=""true"">" &vbcrlf
				  Else
					 XMLStr=XMLStr&" <item name=""typeid"" enabled=""false"">" &vbcrlf
				  End If
			    XMLStr=XMLStr&"  <title>" & request("titletypeid") & "</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>typeid</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>dys</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>0</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>0</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
				  If KS.ChkClng(Request("cbrandid"))=1 Then
					 XMLStr=XMLStr&" <item name=""brandid"" enabled=""true"">" &vbcrlf
				  Else
					 XMLStr=XMLStr&" <item name=""brandid"" enabled=""false"">" &vbcrlf
				  End If
			    XMLStr=XMLStr&"  <title>" & request("titlebrandid") & "</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>brandid</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>dys</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>0</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>0</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
			 End If	
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "select * from ks_field Where (fieldname in(" & AllowField & ") or fieldtype<>0) and ChannelID=" & channelid & " order by orderid,fieldid",conn,1,1
			  If Not RS.Eof Then
					Do While Not RS.Eof
						 If KS.ChkClng(Request("c" & rs("fieldname")))=1 Then
							XMLStr=XMLStr&" <item name=""" & rs("fieldname") &""" enabled=""true"">" &vbcrlf
						 Else
							XMLStr=XMLStr&" <item  name=""" & rs("fieldname") &""" enabled=""false"">" &vbcrlf
						 End If
							XMLStr=XMLStr&"  <title>" & request("title"&rs("fieldname")) & "</title>" &vbcrlf
							XMLStr=XMLStr&"  <fieldtype>" & request("fieldtype"&rs("fieldname")) & "</fieldtype>" &vbcrlf
							XMLStr=XMLStr&"  <fieldname>" & rs("fieldname") & "</fieldname>" &vbcrlf
							XMLStr=XMLStr&"  <condition>" & request("tj"&rs("fieldname")) & "</condition>" &vbcrlf
							XMLStr=XMLStr&"  <showvalue>" & request("xsz"&rs("fieldname")) & "</showvalue>" &vbcrlf
							XMLStr=XMLStr&"  <searchvalue>" & request("ssz"&rs("fieldname")) & "</searchvalue>" &vbcrlf
							XMLStr=XMLStr&" </item>" &vbcrlf
					 RS.MoveNext
					Loop
			  End If
			  
			  '排序字段
			 If KS.ChkClng(Request("corderid"))=1 Then
					XMLStr=XMLStr&" <orderitem name=""id"" enabled=""true"">" &vbcrlf
			 Else
					XMLStr=XMLStr&" <orderitem  name=""id"" enabled=""false"">" &vbcrlf
			 End If
					XMLStr=XMLStr&"  <uptitle>" & request("uptitleid") & "</uptitle>" &vbcrlf
					XMLStr=XMLStr&"  <downtitle>" & request("downtitleid") & "</downtitle>" &vbcrlf
					XMLStr=XMLStr&" </orderitem>" &vbcrlf
					if ChannelID=9 Then
						 If KS.ChkClng(Request("corderadddate"))=1 Then
								XMLStr=XMLStr&" <orderitem name=""date"" enabled=""true"">" &vbcrlf
						 Else
								XMLStr=XMLStr&" <orderitem  name=""date"" enabled=""false"">" &vbcrlf
						 End If
					Else
						 If KS.ChkClng(Request("corderadddate"))=1 Then
								XMLStr=XMLStr&" <orderitem name=""adddate"" enabled=""true"">" &vbcrlf
						 Else
								XMLStr=XMLStr&" <orderitem  name=""adddate"" enabled=""false"">" &vbcrlf
						 End If
					End If
					XMLStr=XMLStr&"  <uptitle>" & request("uptitleadddate") & "</uptitle>" &vbcrlf
					XMLStr=XMLStr&"  <downtitle>" & request("downtitleadddate") & "</downtitle>" &vbcrlf
					XMLStr=XMLStr&" </orderitem>" &vbcrlf
			 If KS.ChkClng(Request("corderhits"))=1 Then
					XMLStr=XMLStr&" <orderitem name=""hits"" enabled=""true"">" &vbcrlf
			 Else
					XMLStr=XMLStr&" <orderitem  name=""hits"" enabled=""false"">" &vbcrlf
			 End If
					XMLStr=XMLStr&"  <uptitle>" & request("uptitlehits") & "</uptitle>" &vbcrlf
					XMLStr=XMLStr&"  <downtitle>" & request("downtitlehits") & "</downtitle>" &vbcrlf
					XMLStr=XMLStr&" </orderitem>" &vbcrlf
			 If KS.ChkClng(Request("cordercmtnum"))=1 Then
					XMLStr=XMLStr&" <orderitem name=""cmtnum"" enabled=""true"">" &vbcrlf
			 Else
					XMLStr=XMLStr&" <orderitem  name=""cmtnum"" enabled=""false"">" &vbcrlf
			 End If
					XMLStr=XMLStr&"  <uptitle>" & request("uptitlecmtnum") & "</uptitle>" &vbcrlf
					XMLStr=XMLStr&"  <downtitle>" & request("downtitlecmtnum") & "</downtitle>" &vbcrlf
					XMLStr=XMLStr&" </orderitem>" &vbcrlf
			  
			  If ChannelID=5 then
				 If KS.ChkClng(Request("cordersalenum"))=1 Then
						XMLStr=XMLStr&" <orderitem name=""salenum"" enabled=""true"">" &vbcrlf
				 Else
						XMLStr=XMLStr&" <orderitem  name=""salenum"" enabled=""false"">" &vbcrlf
				 End If
						XMLStr=XMLStr&"  <uptitle>" & request("uptitlesalenum") & "</uptitle>" &vbcrlf
						XMLStr=XMLStr&"  <downtitle>" & request("downtitlesalenum") & "</downtitle>" &vbcrlf
						XMLStr=XMLStr&" </orderitem>" &vbcrlf
			  End If
			  
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "select * from ks_field Where (fieldtype=4 or fieldtype=12 or fieldtype=5) and ChannelID=" & channelid & " order by orderid,fieldid",conn,1,1
			  If Not RS.Eof Then
					Do While Not RS.Eof
						 If KS.ChkClng(Request("corder" & rs("fieldname")))=1 Then
							XMLStr=XMLStr&" <orderitem name=""" & rs("fieldname") &""" enabled=""true"">" &vbcrlf
						 Else
							XMLStr=XMLStr&" <orderitem  name=""" & rs("fieldname") &""" enabled=""false"">" &vbcrlf
						 End If
							XMLStr=XMLStr&"  <uptitle>" & request("uptitle"&rs("fieldname")) & "</uptitle>" &vbcrlf
							XMLStr=XMLStr&"  <downtitle>" & request("downtitle"&rs("fieldname")) & "</downtitle>" &vbcrlf
							XMLStr=XMLStr&" </orderitem>" &vbcrlf
					 RS.MoveNext
					Loop
			  End If
			'顶部菜单筛选选项
			Dim optionnum:optionnum=KS.ChkClng(Request("optionnum"))
			If optionnum<>0 Then
			  For I=1 To optionnum
			    if request("option"&i)<>"" and request("optionsql"&i)<>"" then
					XMLStr=XMLStr&" <optionitem  name=""" & i &""">" &vbcrlf
					XMLStr=XMLStr&" <title>" & request("option"&i) & "</title>" &vbcrlf
					XMLStr=XMLStr&" <sqlparam><![CDATA[" & request("optionsql"&i) & "]]></sqlparam>" &vbcrlf
					XMLStr=XMLStr&" </optionitem>" &vbcrlf
				end if
			  Next	
			End If
			
			  
		   XMLStr=XMLStr &" </field>" &vbcrlf
		   Call KS.WriteTOFile(KS.Setting(3) & "config/filtersearch/s" & ChannelID & ".xml",xmlstr)
		   RS.Close :Set RS=Nothing
		   KS.Die "<script>top.$.dialog.alert('恭喜，[" & KS.C_S(ChannelID,1) & "]筛选参数配置成功!',function(){location.href='system/KS.Model.asp?Action=SetSearch&ChannelID=" & channelid &"'});</script>"
		 End If
		 	
			set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			FieldXML.async = false
			FieldXML.setProperty "ServerHTTPRequest", true 
			FieldXML.load(Server.MapPath(KS.Setting(3)& "config/filtersearch/s" & ChannelID & ".xml"))
			if FieldXML.parseError.errorCode<>0 Then
				XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
			    XMLStr=XMLStr&"<field>" &vbcrlf
				XMLStr=XMLStr&" <template><![CDATA[{@TemplateDir}/" & KS.C_S(ChannelID,1) & "/筛选模板.html]]></template>" &vbcrlf
				XMLStr=XMLStr&" <isrewrite>0</isrewrite>" &vbcrlf
				XMLStr=XMLStr&" <maxperpage>20</maxperpage>" &vbcrlf
			    XMLStr=XMLStr&" <item name=""tid"" enabled=""true"">" &vbcrlf
			    XMLStr=XMLStr&"  <title>分类</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>tid</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>dy</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>0</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>0</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
				XMLStr=XMLStr&" <orderitem name=""id"" enabled=""true"">" &vbcrlf
				XMLStr=XMLStr&"  <uptitle>按" & KS.C_S(ChannelID,3) & "ID升序</uptitle>" &vbcrlf
				XMLStr=XMLStr&"  <downtitle>按" & KS.C_S(ChannelID,3) & "ID降序</downtitle>" &vbcrlf
				XMLStr=XMLStr&" </orderitem>" &vbcrlf
				XMLStr=XMLStr&" <optionitem  name=""1"">" &vbcrlf
				XMLStr=XMLStr&"  <title>所有" & KS.C_S(ChannelID,3) & "</title>" &vbcrlf
				XMLStr=XMLStr&"  <sqlparam><![CDATA[1=1]]></sqlparam>" &vbcrlf
				XMLStr=XMLStr&" </optionitem>" &vbcrlf
				XMLStr=XMLStr&" <optionitem  name=""2"">" &vbcrlf
				XMLStr=XMLStr&"  <title>推荐" & KS.C_S(ChannelID,3) & "</title>" &vbcrlf
				XMLStr=XMLStr&"  <sqlparam><![CDATA[recommend=1]]></sqlparam>" &vbcrlf
				XMLStr=XMLStr&" </optionitem>" &vbcrlf
			    XMLStr=XMLStr&"</field>" &vbcrlf
                Call KS.WriteTOFile(KS.Setting(3) & "config/filtersearch/s" & ChannelID & ".xml",xmlstr)
			    FieldXML.load(Server.MapPath(KS.Setting(3)& "config/filtersearch/s" & ChannelID & ".xml"))
			End If

		%>
        <script>
		 function CheckForm(){
			  $("#ManageMenuForm").submit();
		 }
		</script>
        <div class="pageCont2">
		 <form name="ManageMenuForm" id="ManageMenuForm" action="KS.Model.asp" method="post">
		 <div class="tabTitle">[<span style='color:red'><%=KS.C_S(ChannelID,1)%></span>]筛选参数设置</div>
         
         <table width='100%' border='0' cellspacing='0' cellpadding='0'>  
		 <%
		  if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("template")
				 If Not Node Is Nothing Then
				  TemplateFile=Node.Text
				 End If
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("isrewrite")
				 If Not Node Is Nothing Then
				  isrewrite=Node.Text
				 End If
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("maxperpage")
				 If Not Node Is Nothing Then
				  maxperpage=Node.Text
				 End If
		  end if
          if KS.IsNul(TemplateFile) Then TemplateFile="{@TemplateDir}/" & KS.C_S(ChannelID,1) & "/筛选模板.html"
		  If KS.IsNul(isrewrite) Then isrewrite=0
		  If KS.IsNul(maxperpage) Then maxperpage=20
		 %>
		 <tr class='tdbg'>
		   <td colspan=10 height='40' style="text-align:left">&nbsp;<strong>绑定模板：</strong><input type='text' name='TemplateFile' id='TemplateFile' class="textbox" value="<%=TemplateFile%>" size="40"/> <%=KSCls.Get_KS_T_C("TemplateFile")%>
		   <%if isrewrite="1" then%>
		   <a href="../../search/c-<%=channelid%>" target="_blank">点此预览</a>
		   <%else%>
		   <a href="../../item/?c-<%=channelid%>" target="_blank">点此预览</a>
		   <%end if%>
		   </td>
		 </tr>
		 <tr class='tdbg'>
		   <td colspan=10 height='40' style="text-align:left">&nbsp;<strong>是否启用伪静态：</strong>
		   <label><input type="radio" name="isrewrite" value="0"<%If isrewrite="0" then response.write " checked"%>/>不开启</label>
		   <label><input type="radio" name="isrewrite" value="1"<%If isrewrite="1" then response.write " checked"%>/>开启(<font color=green>需要服务器支持Rewrite组件</font>)</label>
		   
		   搜索结果每页显示<input type="text" name="maxperpage" class="textbox" value="<%=maxperpage%>" style="text-align:center;width:40px"/>条
		   
		   </td>
		 </tr>
		 <tr class="tdbg">
		   <td style='text-align:center;width:40px' class='sort'>启用</td>
		   <td class='sort'>供选字段</td>
		   <td style='text-align:center' class='sort'>名称</td>
		   <td style='text-align:center' class='sort'>条件</td>
		   <td style='text-align:center' class='sort'>显示的值</td>
		   <td style='text-align:center' class='sort'>搜索的值</td>
		 </tr>
		 <input type="hidden" name="action" value="SetSearch" />
		 <input type="hidden" name="channelid" value="<%=ChannelID%>"/>
		 <input type="hidden" name="flag" value="dosave"/>
		 <%
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@name='tid']")
				 If Not Node Is Nothing Then
				  title=node.selectsinglenode("title").text
				  check=node.selectsinglenode("@enabled").text
				 Else
				  title="分类"
				  check=false
				 End If
		  end if
		 %>
		 <tr class="tdbg" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='utid'>
		  <td class='splittd'  onclick="chk_iddiv('tid')" style='text-align:center' height="30"><input <%if check then response.write " checked"%> onClick="chk_iddiv('tid')" type='checkbox' name='ctid' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('tid')" width="130">所属栏目<span class='tips'>(tid)</span></td>
		  <td class="splittd"><input type="text" name="titletid" size="8" value="<%=title%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		 </tr>
		 <%if KS.ChkClng(KS.C_S(ChannelID,6))=5 then
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@name='price_member']")
				 If Not Node Is Nothing Then
				  title=node.selectsinglenode("title").text
				  check=node.selectsinglenode("@enabled").text
				  xsz=node.selectsinglenode("showvalue").text
				  ssz=node.selectsinglenode("searchvalue").text
				 Else
				  title="商城价":check=false
				  xsz="0-10元,10-100元,100-300元,300-500元,500-1000元,1000元以上"
				  ssz="0-10,10-100,100-300,300-500,500-1000,1000-100000"
				 End If
		  end if
		 
		 %>
		 <tr class="tdbg" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='uprice_member'>
		  <td class='splittd'  onclick="chk_iddiv('price_member')" style='text-align:center' height="30"><input <%if check then response.write " checked"%> onClick="chk_iddiv('price_member')" type='checkbox' name='cprice_member' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('price_member')" width="130">商城价<span class='tips'>(price_member)</span></td>
		  <td class="splittd"><input type="text" name="titleprice_member" size="8" value="<%=title%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'>&nbsp;<select name="tjprice_member">
		  <option value='fw' selected>范围(数字型)</option>
		  </select></td>
		  <td style='text-align:center' class='splittd'><input type="text" style="width:220px" name="xszprice_member" value="<%=xsz%>"  class="textbox"/>
			<div class='tips'>多个用英文逗号隔开如:0-10元,10-100元,100-1000元,1000元以上</div></td>
		  <td style='text-align:center' class='splittd'><input type="text" style="width:220px" name="sszprice_member" value="<%=ssz%>"  class="textbox"/>
			<div class='tips'>多个用英文逗号隔开如:0-10,10-100,100-1000,1000-100000</div></td>
		 </tr>
		 <%
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@name='brandid']")
				 If Not Node Is Nothing Then
				  title=node.selectsinglenode("title").text
				  check=node.selectsinglenode("@enabled").text
				 Else
				  title="品牌":check=false
				 End If
		  end if
		 %>
		 <tr class="tdbg" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='ubrandid'>
		  <td class='splittd'  onclick="chk_iddiv('brandid')" style='text-align:center' height="30"><input <%if check then response.write " checked"%> onClick="chk_iddiv('brandid')" type='checkbox' name='cbrandid' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('brandid')" width="130">所属品牌<span class='tips'>(brandid)</span></td>
		  <td class="splittd"><input type="text" name="titlebrandid" size="8" value="<%=title%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		 </tr>
		 <%elseif KS.ChkClng(KS.C_S(ChannelID,6))=8 then
			 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
					 Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@name='typeid']")
					 If Not Node Is Nothing Then
					  title=node.selectsinglenode("title").text
					  check=node.selectsinglenode("@enabled").text
					 Else
					  title="交易类别":check=false
					 End If
			  end if
		 %>
		 <tr class="tdbg" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='ubrandid'>
		  <td class='splittd'  onclick="chk_iddiv('typeid')" style='text-align:center' height="30"><input <%if check then response.write " checked"%> onClick="chk_iddiv('typeid')" type='checkbox' name='ctypeid' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('typeid')" width="130">交易类别<span class='tips'>(typeid)</span></td>
		  <td class="splittd"><input type="text" name="titletypeid" size="8" value="<%=title%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		 </tr>
		 <%
		 end if
		 
		 
		  set rs=server.CreateObject("adodb.recordset")
		  rs.open "select * from ks_field where (fieldname in(" & AllowField & ") or fieldtype<>0) and channelid=" & channelid & "  order by orderid,fieldid",conn,1,1
		  do while not rs.eof
			if rs("ParentFieldName")="0" or ks.isnul(rs("ParentFieldName"))  then
			if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@name='" & RS("FieldName") &"']")
				 If Not Node Is Nothing Then
				  check=node.selectsinglenode("@enabled").text
				  tj=node.selectsinglenode("condition").text
				  xsz=node.selectsinglenode("showvalue").text
				  ssz=node.selectsinglenode("searchvalue").text
				  title=Node.SelectSingleNode("title").text
				 Else
				  title=rs("title")
				  check=false
				  tj=""
				  xsz=""
				  ssz=""
				 End If
			end if

		  %>
		 <tr class="tdbg" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='u<%=rs("fieldname")%>'>
		 <input type="hidden" name="fieldtype<%=rs("fieldname")%>" value="<%=rs("fieldtype")%>"/>
		  <td class='splittd' onClick="chk_iddiv('<%=rs("fieldname")%>')" style='text-align:center' height="30"><input onClick="chk_iddiv('<%=rs("fieldname")%>')"  type='checkbox' name='c<%=rs("fieldname")%>' value="1"<%if check then response.Write(" checked")%> /></td>
		  <td class='splittd' onClick="chk_iddiv('<%=rs("fieldname")%>')"><%=rs("title")%><span class='tips'>(<%=rs("fieldname")%>)</span></td>
		  <td class='splittd'><input  size="8" type="text" class='textbox' name="title<%=rs("fieldname")%>" value="<%=title%>"/></td>
		  <td style='text-align:center' class='splittd'>
		  <select name="tj<%=rs("fieldname")%>">
		    <%if rs("fieldtype")<>4 and rs("fieldtype")<>12 then%>
		    <option value='dy'<%if tj="dy" then response.write " selected"%>>等于(字符型)</option>
			<%else%>
		    <option value='dys'<%if tj="dys" then response.write " selected"%>>等于(数字型)</option>
		    <option value='fw'<%if tj="fw" then response.write " selected"%>>范围(数字型)</option>
			<%end if%>
			<%if rs("fieldtype")<>3 and rs("fieldtype")<>11 and rs("fieldtype")<>6 and rs("fieldtype")<>7 and rs("fieldtype")<>4 then%>
		    <option value='like'<%if tj="like" then response.write " selected"%>>包含(字符型)</option>
			<%end if%>
		  </select>
		  </td>
		  <td style='text-align:center' class='splittd tips'>
		   <%if rs("fieldtype")=3 or rs("fieldtype")=11 or rs("fieldtype")=6 or rs("fieldtype")=7 then%>
		    <input type="hidden" name="xsz<%=rs("fieldname")%>" value="0">自动显示，字段里设置的选项
		   <%else%>
		    <input type="text" style="width:220px" name="xsz<%=rs("fieldname")%>" value="<%=xsz%>"  class="textbox"/>
			<div class='tips'>多个用英文逗号隔开如:免费,收费,试听</div>
		   <%end if%>
		  </td>
		  <td style='text-align:center' class='splittd tips'>
		  	<%if rs("fieldtype")=3 or rs("fieldtype")=11 or rs("fieldtype")=6 or rs("fieldtype")=7 then%>
		    <input type="hidden" name="ssz<%=rs("fieldname")%>" value="0">自动显示，字段里设置的选项
		   <%else%>
		    <input type="text" style="width:220px" name="ssz<%=rs("fieldname")%>" value="<%=ssz%>"  class="textbox"/>
			<div class='tips'>多个用英文逗号隔开如:0,1,2</div>
		   <%end if%>

		  </td>
		 </tr>
		  <%
		   end if
		   rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		 %>
		 </table>
		<table width='100%' border='0' cellspacing='0' cellpadding='0'> 
		 <tr>
		   <td class="clefttitle" colspan="10" style="text-align:left;height:28px;padding-left:4px;"><strong>排序设置：</strong></td>
		 </tr>
		 <tr class="tdbg">
		   <td style='text-align:center;width:40px' class='sort'>启用</td>
		   <td class='sort'>供选字段</td>
		   <td style='text-align:center' class='sort'>升序名称</td>
		   <td style='text-align:center' class='sort'>降序名称</td>
		 </tr>
		 <%
		 dim uptitle,downtitle
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("orderitem[@name='id']")
				 If Not Node Is Nothing Then
				  uptitle=node.selectsinglenode("uptitle").text
				  downtitle=node.selectsinglenode("downtitle").text
				  check=node.selectsinglenode("@enabled").text
				 Else
				  uptitle="按"& KS.C_S(ChannelID,3) &"ID号升序"
				  downtitle="按"& KS.C_S(ChannelID,3) &"ID号降序"
				  check=false
				 End If
		  end if
		 %>
		 <tr class="tdbg" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='uorderid'>
		  <td class='splittd' style='text-align:center' height="30"><input <%if cbool(check)=true then response.write " checked"%> type='checkbox'  name='corderid' value="1" /></td>
		  <td class='splittd' width="130"><%=KS.C_S(ChannelID,3)%><span class="tips">(id号)</span></td>
		  <td style='text-align:center' class="splittd"><input type="text" name="uptitleid" size="28" value="<%=uptitle%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'><input type="text" name="downtitleid" size="28" value="<%=downtitle%>" class="textbox"/></td>
		 </tr>
		 <%
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("orderitem[@name='adddate']")
				 If Not Node Is Nothing Then
				  uptitle=node.selectsinglenode("uptitle").text
				  downtitle=node.selectsinglenode("downtitle").text
				  check=node.selectsinglenode("@enabled").text
				 Else
				  uptitle="按添加时间升序"
				  downtitle="按添加时间降序"
				  check=false
				 End If
		  end if
		 %>
		 <tr class="tdbg" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='uorderadddate'>
		  <td class='splittd'  style='text-align:center' height="30"><input  <%if cbool(check)=true then response.write " checked"%> type='checkbox'  name='corderadddate' value="1" /></td>
		  <td class='splittd'  width="130">添加时间<span class="tips">(<%if channelid<>9 then response.write "adddate" else response.write "date"%>)</span></td>
		  <td style='text-align:center' class="splittd"><input type="text" name="uptitleadddate" size="28" value="<%=uptitle%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'><input type="text" name="downtitleadddate" size="28" value="<%=downtitle%>" class="textbox"/></td>
		 </tr>
		 
		 <%
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("orderitem[@name='hits']")
				 If Not Node Is Nothing Then
				  uptitle=node.selectsinglenode("uptitle").text
				  downtitle=node.selectsinglenode("downtitle").text
				  check=node.selectsinglenode("@enabled").text
				 Else
				  uptitle="按点击数升序"
				  downtitle="按点击数降序"
				  check=false
				 End If
		  end if
		 %>
		 <tr class="tdbg" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='uorderhits'>
		  <td class='splittd'  style='text-align:center' height="30"><input  <%if cbool(check)=true then response.write " checked"%> type='checkbox'  name='corderhits' value="1" /></td>
		  <td class='splittd'  width="130">浏览数<span class="tips">(hits)</span></td>
		  <td style='text-align:center' class="splittd"><input type="text" name="uptitlehits" size="28" value="<%=uptitle%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'><input type="text" name="downtitlehits" size="28" value="<%=downtitle%>" class="textbox"/></td>
		 </tr>
		 
		 <%
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("orderitem[@name='cmtnum']")
				 If Not Node Is Nothing Then
				  uptitle=node.selectsinglenode("uptitle").text
				  downtitle=node.selectsinglenode("downtitle").text
				  check=node.selectsinglenode("@enabled").text
				 Else
				  uptitle="按评论数升序"
				  downtitle="按评论数降序"
				  check=false
				 End If
		  end if
		 %>
		 <tr class="tdbg" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='uordercmtnum'>
		  <td class='splittd'  style='text-align:center' height="30"><input  <%if cbool(check)=true then response.write " checked"%> type='checkbox'  name='cordercmtnum' value="1" /></td>
		  <td class='splittd'  width="130">评论数<span class="tips">(cmtnum)</span></td>
		  <td style='text-align:center' class="splittd"><input type="text" name="uptitlecmtnum" size="28" value="<%=uptitle%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'><input type="text" name="downtitlecmtnum" size="28" value="<%=downtitle%>" class="textbox"/></td>
		 </tr>
		 <%if channelid=5 then%>
		 <%
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("orderitem[@name='salenum']")
				 If Not Node Is Nothing Then
				  uptitle=node.selectsinglenode("uptitle").text
				  downtitle=node.selectsinglenode("downtitle").text
				  check=node.selectsinglenode("@enabled").text
				 Else
				  uptitle="按销售量升序"
				  downtitle="按销售量降序"
				  check=false
				 End If
		  end if
		 %>
		 <tr class="tdbg" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='uordersalenum'>
		  <td class='splittd'  style='text-align:center' height="30"><input  <%if cbool(check)=true then response.write " checked"%> type='checkbox'  name='cordersalenum' value="1" /></td>
		  <td class='splittd'  width="130">销售量<span class="tips">(salenum)</span></td>
		  <td style='text-align:center' class="splittd"><input type="text" name="uptitlesalenum" size="28" value="<%=uptitle%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'><input type="text" name="downtitlesalenum" size="28" value="<%=downtitle%>" class="textbox"/></td>
		 </tr>
		 <%end if%>
		 
		 <%
		  set rs=server.CreateObject("adodb.recordset")
		  rs.open "select * from ks_field where (fieldtype=4 or fieldtype=12 or fieldtype=5) and channelid=" & channelid & " order by orderid,fieldid",conn,1,1
		  do while not rs.eof
		   	if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("orderitem[@name='" & RS("FieldName") &"']")
				 If Not Node Is Nothing Then
				  check=cbool(node.selectsinglenode("@enabled").text)
				  uptitle=node.selectsinglenode("uptitle").text
				  downtitle=node.selectsinglenode("downtitle").text
				 Else
				  Uptitle="按" & rs("title") & "升序"
				  downtitle="按" & rs("title") & "降序"
				  check=false
				 End If
			end if
		  
		  %>
		 <tr class="tdbg" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='uorder<%=rs("fieldname")%>'>
		  <td class='splittd'  onclick="chk_iddiv('order<%=rs("fieldname")%>')" style='text-align:center' height="30"><input onClick="chk_iddiv('order<%=rs("fieldname")%>')" <%if cbool(check)=true then response.write " checked"%> type='checkbox'  name='corder<%=rs("fieldname")%>' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('order<%=rs("fieldname")%>')" width="130"><%=rs("title")%><span class="tips">(<%=rs("fieldname")%>)</span></td>
		  <td style='text-align:center' class="splittd"><input type="text" name="uptitle<%=rs("fieldname")%>" size="28" value="<%=uptitle%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'><input type="text" name="downtitle<%=rs("fieldname")%>" size="28" value="<%=downtitle%>" class="textbox"/></td>
		 </tr>
		  <%
          rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		 %>
		 </table>
		<table width='100%' border='0' cellspacing='0' cellpadding='0'> 
		 <tr>
		   <td class="clefttitle" colspan="10" style="text-align:left;height:28px;padding-left:4px;"><strong>顶部选项卡筛选项：</strong>  <input type='button' class="button" value="添加一个选项" onClick="doadd(1)"/>
		   </td>
		 </tr>
		 <tr class="tdbg">
		   <td style='padding-left:5px' class='sort' width="20%">选项卡名称</td>
		   <td class='sort' width="80%">SQL查询条件</td>
		 </tr>
		 <tr>
		  <td colspan="2" id="addvote">
		    <table width="100%"  cellpadding="0" cellspacing="0">
			<%
			dim nn,ii
			ii=1
			Set Node=FieldXML.DocumentElement.SelectNodes("optionitem")
			If Node.Length>0 Then
				 For Each nn In FieldXML.DocumentElement.SelectNodes("optionitem")
				 %>
			 <tr class="tdbg">
			   <td style='padding-left:5px' width="20%" class='splittd'><input type='text' name='option<%=ii%>' value="<%=NN.selectSingleNode("title").text%>" class="textbox" />
			  <%if ii=1 then%><div class='tips'>如：推荐信息</div><%end if%>
			   </td>
			   <td class='splittd' width="80%"><input type='text' size="50" name='optionsql<%=ii%>' value="<%=server.HTMLEncode(NN.selectSingleNode("sqlparam").text)%>" class="textbox" />
			   <%if ii=1 then%><div class='tips'>如只需要显示推荐的信息可以输入recommend=1,显示推荐和头条的可以输入 recommend=1 and strip=1等。</div><%end if%>
			   </td>
			 </tr>
				 <%
				 ii=ii+1
				 Next
				 II=II-1
			Else	 
			%>
			 <tr class="tdbg">
			   <td style='padding-left:5px' width="20%" class='splittd'><input type='text' name='option1' value="" class="textbox" />
			   <div class='tips'>如：推荐信息</div>
			   </td>
			   <td class='splittd' width="80%"><input type='text' size="50" name='optionsql1' value="" class="textbox" />
			   <div class='tips'>如只需要显示推荐的信息可以输入recommend=1,显示推荐和头条的可以输入 recommend=1 and strip=1等。</div>
			   </td>
			 </tr>
		 <%end if%>
		 
			 </table>
			 
			 <input type='hidden' name='optionnum'  id='optionnum' value='<%=ii%>'/>

		  </td>
		</tr>
		 <tr class="tdbg">
		   <td colspan=2 style='padding-left:5px'><span class="tips">说明：要删除某个选项，请留空然后保存即可。</span></td>
		 </tr>

		 </table>
		 </form>
      </div>   
	<script type="text/javascript">
    function doadd(num)
    {var i;
    var str="";
    var j=0;
	var optionnum=$("#optionnum").val();
	var id=0;
    for(i=1;i<=num;i++)
    {
	 id=parseInt(optionnum)+i;
     str=str+"<tr class='tdbg'><td style='padding-left:5px' width='20%' class='splittd'><input type=text name=option"+id+" class='textbox'></td><td class='splittd' width=80%><input type=text name='optionsql"+id+"' size=50 class='textbox'></td></tr>";
    }
     jQuery("#addvote").html(jQuery("#addvote").html()+"<table width=100% border=0 cellspacing=0 cellpadding=0>"+str+"</table>");
	 $("#optionnum").val(parseInt(optionnum)+1)
    }
    </script>

		 <div class="attention">
<strong>特别提醒：</strong><br/>
    <li>条件可取 等于（字符型）、等于（数字型）、范围（数字型），包含（字符型）,当指定范围时,还要指定搜索的值,搜索值之间用逗号隔开,如1-20,表示大于1,小于等于20
</li>
	<li>显示的值:即显示在搜索页面供选择的项                     可以带单位,如 1-20万,20-30万</li>
	<li>搜索的值:即用于供数据库搜索的值,搜索值用英文逗号分开      不能带单位,如 1-20,20-30</li>
    
</div>
		 <%
		End Sub
		
		Sub ManageMenu()
		   Dim RS,ChannelID,FieldSql,Doc,Node,XmlFields,XmlFieldArr,Fi,From,xmlname
		   ChannelID=KS.ChkClng(KS.S("ChannelID"))
		   From=KS.S("From")
		   if From="user" then xmlname="usermodelfield" else xmlname="managemodelfield"

		  If KS.G("saveflag")="1" then
		 	set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/" & xmlname & ".xml"))
			Set Node=Doc.documentElement.selectSingleNode("/modelfield/model[@name='" & ChannelID & "']")
			
			 if not node is nothing then  Doc.DocumentElement.RemoveChild(Node)
			 Set Node=Doc.documentElement.appendChild(Doc.createNode(1,"model",""))
			 Node.attributes.setNamedItem(Doc.createNode(2,"name","")).text=channelid
			 Node.text=Replace(KS.S("hasfield")," ","")
			Doc.Save(Server.MapPath(KS.Setting(3)&"Config/" & xmlname & ".xml"))
			Application(KS.SiteSN&"_Config"&xmlname)=empty
			
			if from="" then
			 Call KS.settingsave(2,KS.ChkClng(request("tlen")))  '标题字数
			end if
			
			if request("flag")="menu" then
			 KS.Die "<script>top.box.close();</script>"
			else
			Response.Write "<script>top.$.dialog.alert('恭喜,管理列表菜单配置成功!');</script>"
			end if
          End If
		    XmlFields=LFCls.GetConfigFromXML(xmlname,"/modelfield/model",ChannelID)
			If Not KS.IsNul(XmlFields) Then
			 XmlFieldArr=Split(XmlFields,",")
			End If
		 %>
		 <form name="ManageMenuForm" id="ManageMenuForm" action="KS.Model.asp" method="post">
		 <input type="hidden" name="action" value="ManageMenu" />
		 <input type="hidden" name="channelid" value="<%=ChannelID%>"/>
		 <input type="hidden" name="saveflag" value="1"/>
		 <input type="hidden" name="flag" value="<%=KS.S("flag")%>"/>
		 <input type="hidden" name="from" value="<%=from%>"/>
         <div class="pageCont2">
         <div class="tabTitle">[<span style='color:red'><%=KS.C_S(ChannelID,1)%></span>]模型<%if From="user" then response.write "会员中心" else response.write "后台" %>管理列表菜单设置</div>
		 <table width='100%' border='0' cellspacing='0' cellpadding='0'> 
		 <%if request("flag")<>"menu" then%> 
		 <tr><td height=45 colspan="4">
			 <div class="options">
				  <ul>
				  <li<%If from="user" then response.write " class='curr'"%>><a href="KS.Model.asp?action=ManageMenu&ChannelID=<%=ChannelID%>&from=user">设置会员中心管理列表菜单</a></li>
				  <li<%If from="" then response.write " class='curr'"%>><a href="KS.Model.asp?action=ManageMenu&ChannelID=<%=ChannelID%>">设置后台管理列表菜单</a></li>
				  </ul>
			 </div>
		 </td></tr>
		<%end if%>
		 <tr class="tdbg">
		 <td>
           <style>
		    .list{ padding:0;}
			.list li{float:left;height:25px; width:25%;}
			.list li label{ margin-left:0 !important; margin-right:0 !important;}
		   </style>
          <div class="list clearfix">
          <ul>
          <%
		   Dim FieldsList
		   FieldsList="录入员|Inputer,生成标志|refreshtf,状态|verific,添加时间|AddDate,修改时间|ModifyDate,类型|ModelType,文档属性|Attribute,点击数|Hits,评论数|CmtNum"
		   If channelid<>5 and channelid<>7 and channelid<>8 then 
		    FieldsList=FieldsList&",作者|Author"
		   End If
		   If channelid<>8 then 
		    FieldsList=FieldsList&",等级|Rank"
		   End If
		   If channelid<>5 and channelid<>8 then 
		    FieldsList=FieldsList&",所需费用|ReadPoint"
		   End If
		   FieldsList=FieldsList&",关键字|KeyWords"
		   
		   
		   Select Case KS.ChkClng(KS.C_S(channelid,6))
		    case 1
			  FieldsList=FieldsList&",完整标题|FullTitle,来源|Origin,省份|Province,城市|City"
			case 3
			  FieldsList=FieldsList&",类别|DownLB,语言|DownYY,授权|DownSQ,运行平台|DownPT,演示地址|YSDZ,注册地址|ZCDZ,日下载数|HitsByDay,周下载数|HitsByWeek,月下载数|HitsByMonth"
			case 7
			  FieldsList=FieldsList&",演员|MovieAct,导演|MovieDY,语言|MovieYY,地区|MovieDQ,时长|MovieTime,上映时间|ScreenTime"
		   case 5
			 FieldsList=FieldsList&",单位|Unit,商品编号|Proid,库存量|TotalNum,销售量|SaleNum,参考价|Price,商城价|Price_Member"
		  case 8
			 FieldsList=FieldsList&",价格|price,联系人|ContactMan,电话|Tel,公司|CompanyName,地址|Address,省份|province,城市|City,邮编|Zip,传真|Fax,邮箱|email"
		  end select
		  
		    Set RS = Server.CreateObject("ADODB.RecordSet")
			FieldSql = "SELECT FieldName,Title FROM KS_Field Where fieldtype<>0 and ChannelID=" & ChannelID & " order by orderid asc"
			RS.Open FieldSql, conn, 1, 1
            Do While Not RS.Eof
			 FieldsList=FieldsList&"," & RS("Title") & "|" & RS("FieldName") 
			RS.MoveNext
			Loop
			RS.Close
			Set RS=Nothing
			
			Dim arr:arr=split(FieldsList,",")
			For Fi=0 To Ubound(Arr)
			  response.write "<li><label title='" & arr(FI) &"'><input type=""checkbox"" value=""" &arr(Fi) & """ name=""hasfield"""
			  if instr(lcase(XmlFields),lcase(arr(Fi)))>0 then response.write " checked"
			  response.write ">"&KS.Gottopic(arr(Fi),12) &"</label></li>"
	
			Next
			
			
		   %>
		   
         </ul>
         </div> 
		   

		 </td>
		
		 </tr>
		 <tr class='tdbg'>
		   <td colspan=4 style="height:50px;padding-left:10px; padding-top:0;">
		   <%If request("from")="" then%>
		   标题字数<input type="text" name="Tlen" value="<%=KS.ReadSetting(2)%>" class="textbox" style="text-align:center;width:40px"/>个字
		   <%end if%>
		   <Input type='submit'  value='保存设置' class='button'/></td>
		 </tr>
		 </table>
		 </div>
		 </form>

	<%if request("flag")<>"menu" then%>
		 <div class="attention">
<strong>特别提醒：</strong>
管理列表显示的字段越少则查询显示速度会越快,一般不常用的字段建议不要选择。
</div>
		 <%
	end if
	%>
	<%
		End Sub
 
	
		
		Sub Total()
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From KS_Channel Where ChannelID Not In(6) And ChannelStatus=1  order by channelid asc",conn,1,1
		   With Response
		    .Write "<div class='pageCont2'>"
			.Write "<div class='tabTitle'>各模型信息统计</div>"
           .Write "<dl class=""dtable"">"
		  Do While Not RS.Eof
			.Write "<dd>"
		    If RS("ChannelID")=9 Then
			
		    .Write "<div>" & RS("ChannelName") & "</div>试卷总数：<font color='#ff0000'>" & Conn.Execute("select Count(ID) from KS_SJ")(0) & "</font> 份&nbsp;&nbsp;试题总数：<font color='blue'>" & Conn.Execute("select count(id) from KS_SJTK")(0) & "</font> 题 &nbsp;&nbsp;试题分类：<font color='green'>" & Conn.Execute("select Count(ID) from KS_SJClass")(0) & "</font> 个 "

		   ElseIf RS("ChannelID")=10 Then
		    .Write "<div>" & RS("ChannelName") & "</div>&nbsp;&nbsp;简历总数：<font color='#ff0000'>" & Conn.Execute("select Count(ID) from KS_Job_Resume")(0) & "</font> 个&nbsp;&nbsp;职位总数：<font color='blue'>" & Conn.Execute("select count(id) from KS_Job_zw")(0) & "</font> 个 &nbsp;&nbsp;单位总数：<font color='green'>" & Conn.Execute("select Count(ID) from KS_Job_Company")(0) & "</font> 家 "
		   ElseIf RS("ChannelID")=11 Then
		    .Write "<div>" & RS("ChannelName") & "</div>&nbsp;&nbsp;主题数：<font color='#ff0000'>" & Conn.Execute("select Count(1) from KS_GuestBook")(0) & "</font> 篇&nbsp;&nbsp;论坛版面数：<font color='blue'>" & Conn.Execute("select Count(1) from KS_GuestBoard")(0) & "</font> 个&nbsp;&nbsp;总回复数：<font color='green'>" & Conn.Execute("select sum(PostNum) from KS_GuestBoard")(0) & "</font> 篇"
		   ElseIf RS("ChannelID")=12 Then
		    .Write "<div>" & RS("ChannelName") & "</div>&nbsp;&nbsp;问题数：<font color='#ff0000'>" & Conn.Execute("select Count(1) from KS_AskTopic")(0) & "</font> 个&nbsp;&nbsp;回答数：<font color='blue'>" & Conn.Execute("select Count(1) from KS_AskPosts1")(0) & "</font> 条&nbsp;&nbsp;问答分类：<font color='green'>" & Conn.Execute("select count(1) from KS_AskClass")(0) & "</font> 个"
		   ElseIf RS("ChannelID")=13 Then
		    .Write "<div>" & RS("ChannelName") & "</div>&nbsp;&nbsp;空间总数：<font color='#ff0000'>" & Conn.Execute("select Count(1) from KS_Blog")(0) & "</font> 个&nbsp;&nbsp;企业空间：<font color='blue'>" & Conn.Execute("select Count(1) from KS_EnterPrise")(0) & "</font> 家&nbsp;&nbsp;博文总数：<font color='green'>" & Conn.Execute("select count(1) from KS_BlogInfo")(0) & "</font> 篇"
		   Else
			.Write "<div>" & RS("ChannelName") & "</div>频道总数: <font color=#ff0000>" & Conn.Execute("select count(id) from ks_class where channelid=" & RS("ChannelID") & " and tj=1")(0) & "</font> 个&nbsp;&nbsp;" & RS("ItemName") & "总数: <font color=blue>" & conn.Execute("Select Count(ID) From " & RS("ChannelTable") & " Where DelTF=0")(0) & " </font>" & RS("ItemUnit") & "&nbsp;&nbsp;待审" & RS("ItemName") & ":<font color=green>" & conn.Execute("Select Count(ID) From " & RS("ChannelTable") & " Where  Verific=0")(0) & " </font>" & RS("ItemUnit") & ""
		  End If
			.Write "</dd>"
		    RS.MoveNext
		  Loop
		   .Write "</dl>"
		  End With
		  RS.Close:Set RS=Nothing
		End Sub
		
		'模型设置
		Sub SetChannelParam()
		   With Response
			   Dim ChannelID:ChannelID=KS.ChkClng(KS.G("ChannelID"))
			   If ChannelID=0 Then .Redirect "?": Exit Sub
			   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select * From KS_Channel Where ChannelID=" & ChannelID,Conn,1,3
			   If RS.Eof Then
				 RS.Close:Set RS=Nothing
				.Redirect "?": Exit Sub
			   End If
		     If KS.G("Flag")="ChannelOpenOrClose" Then
			   If RS("ChannelStatus")=1 Then 
				  if conn.execute("select count(channelstatus) from ks_channel where channelstatus=1")(0)=1 then
				   rs.close:set rs=nothing
				   .Write "<script>top.$.dialog.alert('对不起，请至少保持一个模型是开启状态！',function(){ history.back(); });</script>"
				   .end
				   else
					RS("ChannelStatus")=0 
				   End If
			   Else 
			    RS("ChannelStatus")=1
			   end if
			   
			  Dim RSJ,SetArr,SetStr
			 If channelid=10 then
					 Set RSJ=Conn.Execute("Select JobSetting From KS_Config")
					 If Not RSJ.Eof Then
						Dim i,JArr,JobSetting:JobSetting=RSJ(0)
						Jarr=split(JobSetting,"^%^")
						For i=0 To Ubound(jarr)
						  If I=0 Then
							JobSetting=RS("ChannelStatus")
						  Else
							JobSetting=JobSetting & "^%^" & jarr(I)
						  End If
						Next
						Conn.Execute("Update KS_Config Set JobSetting='"& replace(JobSetting,"'","''") &"'")
					 End If
					 RSJ.Close
					 Set RSJ=Nothing
			 ElseIf ChannelID=11 Then '论坛
			         Set RSJ=Conn.Execute("Select Setting From KS_Config")
					 If Not RSJ.Eof Then
						  SetArr=Split(RSJ("Setting")&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","^%^")
						  For I=0 To Ubound(SetArr)
						   If I=0 Then 
							SetStr=SetArr(0)
						   ElseIf I=56 Then
							SetStr=SetStr & "^%^" & RS("ChannelStatus")
						   Else
							SetStr=SetStr & "^%^" & SetArr(I)
						   End If
						  Next
						  Conn.Execute("Update KS_Config Set Setting='"& replace(SetStr,"'","''") &"'")
					End If
			         RSJ.Close
					 Set RSJ=Nothing
			 ElseIf ChannelID=12 Then '问答
			         Set RSJ=Conn.Execute("Select AskSetting From KS_Config")
					 If Not RSJ.Eof Then
						  SetArr=Split(RSJ("AskSetting")&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","^%^")
						  For I=0 To Ubound(SetArr)
						   If I=0 Then 
							SetStr=RS("ChannelStatus")
						   Else
							SetStr=SetStr & "^%^" & SetArr(I)
						   End If
						  Next
						  Conn.Execute("Update KS_Config Set AskSetting='"& replace(SetStr,"'","''") &"'")
					End If
			         RSJ.Close
					 Set RSJ=Nothing
			  End If
			 
			 End If
			 RS.Update
			 RS.Close:Set RS=Nothing
			  Call KS.DelCahe(KS.SiteSn & "_Config")
			   Call KS.DelCahe(KS.SiteSN & "_selectallowclass")
				 Call KS.DelCahe(KS.SiteSN & "_selectclass")
				 Call KS.DelCahe(KS.SiteSN & "_classpath")
				 Call KS.DelCahe(KS.SiteSN & "_classnamepath")
				 Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
			  
			 Session("FromFile")="System/KS.Model.asp"
			 .Write "<script>top.location.href='../index.asp';</script>"
		   End With
		End Sub
		
		Sub Main()
		Response.Write ""&_
		"<table width='100%' border='0' cellspacing='0' cellpadding='0'>"&_
		" <tr><td height=5></td></tr>"&_
		"</table></div>" 
		Response.Write "<script>"
		Response.Write "$(document).ready(function(){"
		Response.Write "parent.frames['BottomFrame'].Button1.disabled=true;"
		Response.Write "parent.frames['BottomFrame'].Button2.disabled=true;"
		Response.Write "})</script>"
		%>
        <script type="text/javascript">
		 function delPlusl(obj,Pluslid){
		   top.$.dialog.confirm('是否关闭模块?',function(){ location.href='<%=KS.Setting(3) & KS.Setting(89)%>system/KS.Model.asp?action=SetChannelParam&flag=ChannelOpenOrClose&channelid='+Pluslid;},function(){});
		 }
		 function addPlusl(obj,Pluslid){
		  top.$.dialog.confirm('是否开启模块?',function(){location.href='<%=KS.Setting(3) & KS.Setting(89)%>system/KS.Model.asp?action=SetChannelParam&flag=ChannelOpenOrClose&channelid='+Pluslid;},function(){});
		 }
		 function delPluslapp(obj,Pluslid){
		    top.$.dialog.confirm('是否关闭应用插件?',function(){location.href='<%=KS.Setting(3) & KS.Setting(89)%>system/KS.Model.asp?action=AddC&action1=delPluslapp&Pluslid='+Pluslid;},function(){});
		 }
		 function addPluslapp(obj,Pluslid){
		    top.$.dialog.confirm('是否开启应用插件?',function(){location.href='<%=KS.Setting(3) & KS.Setting(89)%>system/KS.Model.asp?action=AddC&action1=addPluslapp&Pluslid='+Pluslid;},function(){});
		 }
		</script>
        <%

		if KS.G("action1")="delPlusl" then
				SetChannelParam
		end if
		if KS.G("action1")="addPlusl" then
				SetChannelParam
		end if
		
		if KS.G("action1")="delPluslapp" then
			if ks.CheckFile("../plus/plus_" & KS.G("Pluslid")&"/Config.xml") then 
				call EditXMLid("../plus/plus_" & KS.G("Pluslid")&"/Config.xml","0")
				 Session("FromFile")="System/KS.Model.asp"
				 KS.Die("<script>top.$.dialog.alert('关闭应用插件成功!',function(){top.location.href='index.asp?from=app';});</script>")
			end if
		end if
		if KS.G("action1")="addPluslapp" then
			if ks.CheckFile("../plus/plus_" & KS.G("Pluslid")&"/Config.xml") then 
				call EditXMLid("../plus/plus_" & KS.G("Pluslid")&"/Config.xml","1")
				 Session("FromFile")="System/KS.Model.asp"
				 Response.Write("<script>top.$.dialog.alert('开启应用插件成功!',function(){top.location.href='index.asp?from=app';});</script>")
			end if
		end if

		%>
        
		<div class="Pluslist">
        <div class="tabTitle" style="padding-left:0;">系统模型 <a href="javascript:;" onClick="location.href='?action=Add';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Go&OpStr=<%=Server.URLEncode("模型管理 >> <font color=red>添加模型</font>")%>';" style="font-size: 12px;color: #247ec0;">添加模型</a></div>
        <div class="co">
        <ul>
        <%
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	    RS.Open "Select * From KS_Channel where channelid  not in(" & ChannelNotOnStr &") Order By orderid,ChannelID",conn,1,1
		Do While Not RS.Eof
		 Dim ModelEname:ModelEname = KS.C_S(RS("BasicType"),10)
		 If KS.CheckFile("../" &ModelEname&"/Config.xml") Then
		  if RS("ChannelStatus")=1 then %>
        	<li id="chover" class="clearfix">
			<div class="left">
				<i class="icon model<%=rs("basictype")%>"></i>
				<br>
				<i class="icon no"></i>
			</div>
			<div class="right">
			<%
			if rs("channelid")=10 Then
			 Response.Write "<a title='参数配置' href='../job/KS.JobSetting.asp?from=model' onclick=""$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=GoSave&OpStr=" & server.URLEncode("模型管理 >> <font color=red>求职系统参数设置</font>") &"';"">"& RS("ChannelName")&"</a> "
			Elseif rs("channelid")=11 Then
			 Response.Write "<a title='参数配置' href='../club/KS.GuestSetting.asp?from=model' onclick=""$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=GoSave&OpStr=" & server.URLEncode("模型管理 >> <font color=red>论坛参数设置</font>") &"';"">"& RS("ChannelName")&"</a> "
			ElseIf rs("ChannelID")=12 Then 
			 Response.Write "<a title='参数配置' href='../ask/KS.AskSetting.asp?from=model' onclick=""$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=GoSave&OpStr=" & server.URLEncode("模型管理 >> <font color=red>问答参数设置</font>") & "';"">"& RS("ChannelName")&"</a> "
			ElseIf rs("ChannelID")=13 Then 
			 Response.Write "<a title='参数配置' href='../space/KS.SpaceSetting.asp?from=model' onclick=""$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=GoSave&OpStr=" & server.URLEncode("模型管理 >> <font color=red>空间参数设置</font>") & "';"">"& RS("ChannelName")&"</a> "
			Elseif rs("channelid")=1 or (Instr(channelNotOnStr,rs("channelid"))=0 and rs("channelid")<>10) then
			 Response.Write "<a title='参数配置' href='?action=Edit&ChannelID=" & rs("ChannelID") & "' onclick=""$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=GoSave&OpStr=" & server.URLEncode("模型管理 >> <font color=red>修改模型配置</font>") &"';"">"& RS("ChannelName")&"</a> "
			else
			 Response.Write "<font title='参数配置' color=#a7a7a7>"& RS("ChannelName")&"</font> "
			end if
            %>
			<!--<img src="../images/check.png" width="16" height="16" />--><div class="CPlus" id="sub<%=rs("ChannelID")%>"><a  href="javascript:void(0)" onClick="delPlusl(this,<%=RS("ChannelID")%>)">关闭</a> 
            <%
			
			 If RS("ChannelID")>=100 Then
			 Response.Write "<a href='?action=Del&ChannelID=" & rs("ChannelID") & "' onclick='return(confirm(""此操作不可逆，确定删除吗？""))'>删除</a> "
			 End If
			 
			 IF KS.ChkClng(rs("BasicType"))<11 And (rs("ChannelID")<>6 and rs("channelid")<>10  and Instr(channelNotOnStr,rs("channelid"))=0) then
			 Response.Write "<a href='KS.Field.asp?ChannelID=" & rs("ChannelID") & "' onClick=""$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=GoSave&OpStr=" & server.URLEncode("模型管理 >> <font color=red>模型字段管理</font>") &"';"">字段</a> "
			  Response.Write " <a href='?Action=SetSearch&ChannelID=" & RS("ChannelID") & "' onClick=""$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=SetParam&OpStr=" & server.URLEncode("模型管理 >> <font color=red>设置筛选</font>") &"';"">筛选</a>"
			 end if
			
			%>
            
              </div>
            </div>
			</li>
        <%else%>
        	<li>
				<div class="left">
					<i class="icon nomodel<%=rs("basictype")%>"></i>
					<br>
					<i class="icon yes"></i>
				</div>
				<div class="right">
				<%=RS("ChannelName")%><!--<img src="../images/check1.png"  width="16" height="16"/>--><div id="sub<%=rs("ChannelID")%>" class="sub"><a  href="javascript:void(0)" onClick="addPlusl(this,<%=RS("ChannelID")%>)" class="open">开启</a></div>
				</div>
			</li>
        <%end if
	    Else
		   conn.execute("update ks_channel set ChannelStatus=0 where channelid=" & rs("channelid"))
		End If
		RS.MoveNext
		Loop
		RS.Close
		Set RS=Nothing	
		%>
        
       
        </ul>
        </div>
        </div>
        <div class="clear blank10" ></div>
        <a name="app"></a>
        <div class="Pluslist">
        <div class="tabTitle" style="padding-left:0;">应用插件</div>
        <div class="co">
        <ul>
        <%
		Dim FsoItem,APPXMLStr
		Dim FsoObj:Set FsoObj = KS.InitialObject(KS.Setting(99))
		Dim FolderObj:Set FolderObj = FsoObj.GetFolder(Server.MapPath("../plus"))
		Dim SubFolderObj:Set SubFolderObj = FolderObj.SubFolders
		
		For Each FsoItem In SubFolderObj
		   if KS.CheckFile("../plus/"&FsoItem.name&"/Config.xml") then
		       Dim FieldXML:set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			   FieldXML.async = false
			   FieldXML.setProperty "ServerHTTPRequest", true 
			   FieldXML.load(Server.MapPath("../plus/"&FsoItem.name&"/Config.xml"))
			   if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
			    Dim NodeXML:Set NodeXML=FieldXML.DocumentElement.SelectNodes("App")		
			    If NodeXML.length>0 Then
			    Dim AppName:AppName=FieldXML.DocumentElement.SelectSingleNode("App/AppName").Text
			    Dim AppStatus:AppStatus=FieldXML.DocumentElement.SelectSingleNode("App/AppStatus").Text
				Dim AppEname:AppEname=Replace(lcase(FsoItem.name),"plus_","")
				
				
				APPXMLStr=APPXMLStr & "   <app name=""" & AppEname &""" status=""" & AppStatus & """><![CDATA[" & AppName & "]]></app>" &vbcrlf
				
				
				
		        If AppStatus="0" Then
				 %>
					<li id="chover">
					<div class="left1">
						<i class="icon app<%=AppEname%>"></i>
						<br/>
						<i class="icon yes"></i>
					</div>
			        <div class="right1">
					  <div class="check1">
					   <!--<img src="../images/check1.png" width="16" height="16"/>--><%=AppName%></div>
					  <div id="sub<%=AppEname%>" class="sub"><a  href="javascript:void(0)" onClick="addPluslapp(this,'<%=AppEname%>')">开启</a></div>
					</div>
					</li>
				<%
			    Else
				 %>
					<li id="chover">
					<div class="left1">
						<i class="icon noapp<%=AppEname%>"></i>
						<br>
						 <i class="icon no"></i>
					 </div>	 
					 <div class="right1">
					  <div class="check"><%=AppName%></div>
					  <!--<img src="../images/check.png"  width="16" height="16"/>-->
					 <div class="CPlus" class="sub" id="sub<%=AppEname%>"><a  href="javascript:void(0)" onClick="delPluslapp(this,'<%=AppEname%>')">关闭</a></div>
                     </div>		
					</li>
				<%
				End If
				
				
			    End If
			  End If
		  End If
		Next
		
		
		APPXMLStr="<?xml version=""1.0"" encoding=""utf-8""?>" & vbcrlf &" <MyApp>" & vbcrlf &  APPXMLStr &" </MyApp>" &vbcrlf
		If  KS.WriteTOFile(KS.Setting(3) & "config/AppSetting.xml",APPXMLStr)=false Then
		   KS.Die ("<script>alert('提示：插件配置文件/config/AppSetting.xml写入失败，请检查config目录是否有写入权限！');</script>")
		End If
		
		%>
        </ul>
        </div>
        </div>
		<%		
		End Sub
		
		Sub EditXMLid(XMLurl,setc)
			dim XMLStr,FieldXML,Nodek,NodeXML,Fast,Fasturl,Attribute,Role,Fastico,Nodek2,NodeXML2,Mchannelid
			set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			FieldXML.async = false
			FieldXML.setProperty "ServerHTTPRequest", true 
			FieldXML.load(Server.MapPath(XMLurl))
			Set NodeXML=FieldXML.DocumentElement.SelectSingleNode("App")	
			If Not NodeXML Is Nothing Then
			   NodeXML.SelectSingleNode("AppStatus").Text=setc
			   FieldXML.Save(Server.MapPath(XMLurl))
			End If
			
		End Sub
		
		
		Sub ChannelAddOrEdit()
		Dim SqlStr, RS, InstallDir, FsoIndexFile,StaticTF, FsoIndexExt,FsoListNum,i,ThumbnailsConfig
		Dim ChannelName,ModelEname,ChannelTable,ChannelStatus,WapSwitch,WapSearchTemplate,ItemName,ItemUnit,FsoFolder,Descript,ModelIco,ModelShortName,CommentTime
		Dim FsoHtmlTF,BasicType,MaxPerPage,UserTF,UserClassStyle,UserEditTF
		Dim UpFilesTF,UpfilesDir,UserUpFilesTF,UserUpfilesDir,UserSelectFilesTF,UpfilesSize,AllowUpPhotoType,AllowUpFlashType,AllowUpMediaType,AllowUpRealType,AllowUpOtherType
		Dim  UserAddMoney,UserAddPoint,UserAddScore,RefreshFlag,InfoVerificTF,VerificCommentTF,CommentVF,CommentLen,CommentTemplate,ChargeType,DiggByVisitor,DiggByIP,DiggRepeat,DiggPerTimes
		Dim FsoContentRule,FsoClassListRule,FsoClassPreTag,LatestNewDay,PubTimeLimit,AnnexPoint,SJzsdRule,arrGroupID'ZSD_开发
		Dim ChannelID:ChannelID = KS.ChkClng(KS.G("ChannelID"))
		
	'	On Error Resume Next
	   If KS.G("Action")="Edit" Then
			SqlStr = "select * from KS_Channel Where ChannelID=" & ChannelID
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1,1
			ChannelName   = RS("ChannelName")
			ModelEname    = RS("ModelEname")
			ChannelTable  = RS("ChannelTable")
			ItemName      = RS("ItemName")
			ItemUnit      = RS("ItemUnit")
			ChannelStatus = RS("ChannelStatus")
			StaticTF      = RS("StaticTF")
			FsoFolder     = RS("FsoFolder")
			FsoListNum    = RS("FsoListNum")
			WapSwitch     = RS("WapSwitch")
			Descript      = RS("Descript")
			BasicType     = RS("BasicType")
			UserTF        = RS("UserTF")
			UserClassStyle= RS("UserClassStyle")
			UserEditTF    = RS("UserEditTF")
			FsoHtmlTF     = RS("FsoHtmlTF")
			UpFilesTF     = RS("UpFilesTF")
			UpfilesDir    = RS("UpfilesDir")
			UserUpFilesTF = RS("UserUpFilesTF")
			UserUpfilesDir= RS("UserUpfilesDir")
			UserSelectFilesTF =RS("UserSelectFilesTF")
			UpfilesSize   = RS("UpfilesSize")
			AllowUpPhotoType = RS("AllowUpPhotoType")
			AllowUpFlashType = RS("AllowUpFlashType")
			AllowUpMediaType = RS("AllowUpMediaType")
			AllowUpRealType  = RS("AllowUpRealType")
			AllowUpOtherType = RS("AllowUpOtherType")
			ThumbnailsConfig = RS("ThumbnailsConfig")&"|0|||||||||||||"
			
			UserAddMoney     = RS("UserAddMoney")
			UserAddPoint     = RS("UserAddPoint")
			UserAddScore     = RS("UserAddScore")
			RefreshFlag      = RS("RefreshFlag")
			MaxPerPage       = RS("MaxPerPage")
			InfoVerificTF    = RS("InfoVerificTF")
			VerificCommentTF = RS("VerificCommentTF")
			CommentVF        = RS("CommentVF")
			CommentLen       = RS("CommentLen")
			CommentTemplate  = RS("CommentTemplate")
			WapSearchTemplate= RS("WapSearchTemplate")
			ChargeType       = RS("ChargeType")
			AnnexPoint       = RS("AnnexPoint")
			DiggByVisitor    = RS("DiggByVisitor")
			DiggByIP         = RS("DiggByIP")
			DiggRepeat       = RS("DiggRepeat")
			DiggPerTimes     = RS("DiggPerTimes")
			FsoContentRule   = RS("FsoContentRule")
			FsoClassListRule = RS("FsoClassListRule")
			FsoClassPreTag   = RS("FsoClassPreTag")
			LatestNewDay     = RS("LatestNewDay")
			PubTimeLimit     = RS("PubTimeLimit")
			ModelShortName   = RS("ModelShortName")
			ModelIco         = RS("ModelIco")
		Else
			  ChannelStatus =1 
			  ItemName="文章"
			  ItemUnit="篇"
			  ThumbnailsConfig="0.3|130|90|1|0|0|||||||||||||||||5|10|2|0|栏目||||||||||||"
			  UpfilesDir  = "UploadFiles/"
			  UserUpfilesDir = "User/"
			  UserAddMoney=0
			  UserAddPoint=0
			  UserAddScore=0
			  CommentLen=0
			  ModelShortName="文章"
			  UpfilesSize = 1024
			  BasicType   = 1
			  MaxPerPage=20
			  FsoFolder="html/"
			  RefreshFlag=2
			  InfoVerificTF=1
			  VerificCommentTF=0
			  UserTF=1
			  UserEditTF=0
			  UserClassStyle=1
			  UpFilesTF=1
			  AllowUpPhotoType = "gif|jpg|png"
			  AllowUpFlashType = "swf"
			  AllowUpMediaType = "mid|mp3|wmv|asf|avi|mpg"
			  AllowUpRealType  = "ram|rm|ra"
			  AllowUpOtherType = "rar|doc|zip"
			  WapSwitch = 1
			  ChargeType=1
			  FsoListNum=3
			  DiggByVisitor    = 0
			  DiggRepeat       = 0
			  DiggPerTimes     = 1
			  FsoClassPreTag="list"
			  FsoClassListRule = "1"
			  FsoContentRule   = "{$ClassEname}_{$ClassID}_"
			  LatestNewDay     = 3 
			  PubTimeLimit     = 20
			  AnnexPoint       = 0
			  ModelIco         = "/user/images/icon13.png"
			  SJzsdRule        = "1" 
			  CommentTemplate  = "{@TemplateDir}/文章系统/评论页.html"
		End If
			  ThumbnailsConfig=Split(ThumbnailsConfig&"||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||","|")
			  If Ubound(ThumbnailsConfig)<2 Then
			   ThumbnailsConfig(0)=0.3
			   ThumbnailsConfig(1)=130
			   ThumbnailsConfig(2)=90
			   ThumbnailsConfig(3)=0
			   ThumbnailsConfig(4)=0
			  End IF
		With Response
		.Write ""&_
		"<title>模型基本参数设置</title>" &_
		"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" &_
		"<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"&_
		"<script src=""../../KS_Inc/JQuery.js"" language=""JavaScript""></script>"&_
		"<script src=""../images/pannel/tabpane.js"" language=""JavaScript""></script>" & _
		"<link href=""../images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & _
		"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"&_
		"<body>" &_
		"<div class='tabTitle'>网站模型管理</div><div class='tab-page tab-page2' id='modelpanel'>"& _
		"<form id=""myform"" name=""myform"" method=""post"" action=""KS.Model.asp?Action=EditSave&ChannelID=" & ChannelID & """ >" & _
        " <SCRIPT type=text/javascript>"& _
        "   var tabPane1 = new WebFXTabPane( document.getElementById( ""modelpanel"" ), 1 )"& _
        " </SCRIPT>"& _
		" <div class=tab-page id=site-page>"& _
		"  <H2 class=tab>基本信息</H2>"& _
		"	<SCRIPT type=text/javascript>"& _
		"				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"& _
		"	</SCRIPT>" & _
		"<dl class=""dtable"">"
		.Write "    <dd><div>模型状态：</div><input type=""radio"" name=""ChannelStatus"" value=""1"" "
		If ChannelStatus = 1 Then .Write (" checked")
		.Write ">"
		.Write "正常"
		.Write "  <input type=""radio"" name=""ChannelStatus"" value=""0"" "
		If ChannelStatus = 0 Then .Write (" checked")
		.Write ">"
		.Write "关闭<span>只有设置“正常”的模型才可以正常使用。</span></dd>"
		.Write "    <dd><div>手机版状态：</div><input type=""radio"" name=""WapSwitch"" value=""1"" "
		If WapSwitch = 1 Then .Write (" checked")
		.Write ">"
		.Write "正常"
		.Write "  <input type=""radio"" name=""WapSwitch"" value=""0"" "
		If WapSwitch = 0 Then .Write (" checked")
		.Write ">"
		.Write "关闭"
		.Write "   <span>只有设置“正常”的模型才可以正常使用。</span> </dd>"
%>
		<script type="text/javascript">
		 $(document).ready(function(){
		   $("input[name=FsoHtmlTF]").click(function(){
		     FsoDisplay();
			});
		   FsoDisplay();
		 });
		 function FsoDisplay()
		 {
		   var FsoHtmlTF=$("input[name=FsoHtmlTF]:checked").val();
		   if (FsoHtmlTF==0){
		    $("#fsoarea").hide();
			$("#staticarea").show();
		   }else if(FsoHtmlTF==1){
		    $("#fsoarea").show();
		    $("#staticarea").hide();
		   }else{
		    $("#fsoarea").show();
			$("#staticarea").show();
		   }
		 }
	
		 function CheckForm()
		 {  
		  if ($("input[name=ChannelName]").val()=="")
		  {
		   top.$.dialog.alert('请输入模型名称',function(){$("input[name=ChannelName]").focus();});
		   return false;
		  }
		  if ($("input[name=ModelEname]").val()=="")
		  {
		   top.$.dialog.alert('请输入模型的目录名称',function(){$("input[name=ModelEname]").focus()});
		   return false;
		  }
		  if ($("input[name=ChannelTable]").val()=="")
		  {
			top.$.dialog.alert('请输入数据名！',function(){$("input[name=ChannelTable]").focus()});
			 return false;
		  }
		  if ($("input[name=ItemName]").val()=="")
		  {
			 top.$.dialog.alert('请输入项目名称！',function(){$("input[name=ItemName]").focus()});
			 return false;
		  }
		  if ($("input[name=ItemUnit]").val()=="")
		  {
			 top.$.dialog.alert('请输入项目单位！',function(){$("input[name=ItemUnit]").focus();});
			 return false;
		  }
		  if ($("input[name=FsoFolder]").val()=="")
		  {
			 top.$.dialog.alert('请输入模型目录！',function(){$("input[name=FsoFolder]").focus();});
			 return false;
		  }
		  $("#myform").submit();
		 }
		 function GetTable(val)
		 { 
		    $.get('../../plus/ajaxs.asp', { foldername: escape($('input[name=ChannelName]').val()), action: 'Ctoe' },function(data){
			$('input[name=ChannelTable]').val(unescape(data));
		    $('input[name=ModelEname]').val(unescape(data));
		    $('input[name=FsoFolder]').val('html/'+data+'/');
		    $('input[name=UserUpfilesDir]').val(data+'/');
		    $('input[name=UpfilesDir]').val('UploadFiles/'+data+'/');
			 });
		 }
		</script>
		<style type="text/css">
		 .textbox{border:0px;border-bottom:1px solid #000;width:60px;background:transparent }
		.tips {color: #999999;padding:2px}
		.txt {color: #666;border:1px solid #ccc;height:22px;line-height:22px; width:200px;margin-bottom:0 !important;}
		textarea {color: #666;border:1px solid #ccc;}
		</style>
		
		<dd><div><strong>模型名称：</strong></div><input class="txt textbox" name="ChannelName" type="text" <%If KS.G("Action")<>"Edit" Then Response.Write " onkeyup='GetTable(this.value)'"%> value="<%=ChannelName%>" size="50">
        <span>如：文章系统，图片系统等。</span>
		</dd>
		<dd><div><strong>模型目录：</strong></div><input class="txt textbox" name="ModelEname" type="text"<%If KS.G("Action")="Edit" Then Response.Write " Disabled"%> value="<%=ModelEname%>" size="50"> <span class="tips">只能用字母和数字的组合，且不能修改</span>
		</dd>

		<dd><div><strong>数据表名称：</strong></div><%If KS.G("Action")="Add" Then Response.Write " KS_U_" %><input name="ChannelTable" id='ChannelTable' type="text" value="<%=ChannelTable%>" class="txt textbox" size="50"<%If KS.G("Action")="Edit" Then Response.Write " Disabled"%>><span>说明：创建数据表后无法修改，并且用户创建的数据表以"KS_U_"开头</span> 
		</dd> 
		<dd><div><strong>基 类 型：</strong></div><select name="BasicType" id="BasicType" <%If KS.G("Action")="Edit" Then Response.Write " Disabled"%>>
			 <option value=1<%if BasicType="1" Then Response.Write " selected"%>>文章类型</option>
			 <option value=2<%if BasicType="2" Then Response.Write " selected"%>>图片类型</option>
			 <option value=3<%if BasicType="3" Then Response.Write " selected"%>>软件类型</option>
			 <%if instr(ChannelNotOnStr,"4")=0 then%>
			<option value=4<%if BasicType="4" Then Response.Write " selected"%>>动漫类型</option>
			 <%end if%>
			 <%If KS.G("Action")="Edit" Then%>
			 <option value=5<%if BasicType="5" Then Response.Write " selected"%>>商城类型</option>
			 <option value=6<%if BasicType="6" Then Response.Write " selected"%>>音乐类型</option>
			 <option value=7<%if BasicType="7" Then Response.Write " selected"%>>影视类型</option>
			 <option value=8<%if BasicType="8" Then Response.Write " selected"%>>供求类型</option>
			 <option value=9<%if BasicType="9" Then Response.Write " selected"%>>考试类型</option>
			 <%End If%>
			</select>
			</dd>
		<dd><div><strong>项目名称：</strong></div><input class="txt textbox" name="ItemName" id="ItemName" type="text" value="<%=ItemName%>" size="50"> <span class="tips">*如：文章、图片、软件等项</span></dd>
		<dd><div><strong>项目单位：</strong></div><input name="ItemUnit" type="text" value="<%=ItemUnit%>" class="txt textbox" size="50"> <span class="tips">*如：篇、个、本等</span></dd>
		<dd><div><strong>栏目名称：</strong></div><input name="FolderName" type="text" value="<%=ThumbnailsConfig(26)%>" class="txt textbox" size="50"> <span class="tips">*如：分类，品牌等</span></dd>
		
		<dd><div><strong>模型描述：</strong></div><textarea name="Descript" rows=4 cols=80 class="ml10"><%=Descript%></textarea></dd>
		</dl>
		</div>
		
		
		<%
		If ChannelID=9 Then
		.Write " <div class=tab-page id=fso-page>"
		.Write "  <H2 class=tab>系统选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""fso-page"" ) );"
		.Write "	</SCRIPT>"

		.Write "<dl class=""dtable"">"
	    .Write "<dd><div><strong>生成的总目录：</strong></div><input class='txt' name='FsoFolder' type='text' value='" & FsoFolder & "' size='20'><span class='tips'>*用于生成静态html存放的目录，只能以字母和数字的组合,必须以""/""结束</span>"
		.Write "</dd>"
		
		.Write "<dd><div><strong>考后心得管理：</strong></div>"
		.Write " <strong>发布心得:</strong> <input type=""radio""  name=""ksxdsh"" value=""1"" "
			If ThumbnailsConfig(19) = "1" Then .Write (" checked")
			.Write ">不用审核"
	
			.Write "<input type=""radio""  name=""ksxdsh"" value=""0"" "
			If ThumbnailsConfig(19) = "0" Then .Write (" checked")
			.Write ">"
			.Write "需要审核<br/>"
		.Write "  <strong>发布限制:</strong> <input type=""radio""  name=""ksxd"" value=""1"" "
			If ThumbnailsConfig(18) = "1" Then .Write (" checked")
			.Write ">无限制"
	
			.Write "<input type=""radio""  name=""ksxd"" value=""0"" "
			If ThumbnailsConfig(18) = "0" Then .Write (" checked")
			.Write ">"
			.Write "有考试过才可以提交心得<br/>"
		.Write "</dd> "
		
		Dim ModelChargeType:ModelChargeType=KS.ChkClng(KS.C_S(9,34))
		Dim ChargeUnit
		If ModelChargeType=0 Then 
		  ChargeUnit=KS.Setting(46)&KS.Setting(45)
		ElseIf ModelChargeType=1 Then
		  ChargeUnit="元"
		Else
		  ChargeUnit="积分"
		End If
		
	    .Write "<dd><div><strong>日常练习免费练习次数<font>(有效期用户每天练习次数不受限制)</font>：</strong></div><input class='txt' style='text-align:center' name='FsoListNum' type='text' value='" & FsoListNum & "' size='6'>次</dd> "
	    .Write "<dd><div><strong>日常练习收费设置<font>(有效期用户除外)</font>：</strong></div>超过上面设置的免费次数将按<input class='txt' style='text-align:center' name='FsoClassListRule' type='text' value='" &FsoClassListRule & "' size='3'>" & ChargeUnit &"/次收取</dd> "
	    .Write "<dd><div><strong>日常练习随机抽题数：</strong></div>"      
		.Write " <input maxlength='3' class='txt' style='text-align:center' name='FsoHtmlTF' type='text' value='" & FsoHtmlTF & "' size='6'> 题 每次练习最高奖励积分<input maxlength='3' class='txt' style='text-align:center' name='rclxjfjl' type='text' value='" & ThumbnailsConfig(14) & "' size='6'>分。<br/><span>Tips:积分赠送的计算公式：（答对题/实际抽取的总题数）X 最高积分,再四舍五入取整，小于1分将不赠送。<Br/>如：随机抽取10道题,最高奖励2分，那么答对5题的情况，将奖励1分。<br/></span> "
		.Write "</dd>"
	    .Write "<dd><div><strong>答题鼓励语言设置：</strong></div>"      
		.Write " 正确：<div class=""clear""></div><textarea class='txt' style='width:260px;height:100px' name='rightyy'>" & ThumbnailsConfig(9)& "</textarea>表情目录：<input class='txt' name='imgFsoFolder1' type='text' value='" & ThumbnailsConfig(20) & "' size='20'> 输入0不启用.[只支持.gif图片格式]<br/>一行一个鼓励语言。<div class=""clear""></div><br/><br/>错误：<div class=""clear""></div><textarea class='txt' style='width:260px;height:100px' name='wrongyy'>" & ThumbnailsConfig(10)& "</textarea>表情目录：<input class='txt' name='imgFsoFolder2' type='text' value='" & ThumbnailsConfig(21) & "' size='20'> 输入0不启用<br/>一行一个鼓励语言。"
		.Write "</dd>"
	    .Write "<dd><div><strong>启用考试成绩积分奖励设置：</strong></div>"
			.Write "  <input type=""radio"" onclick=""$('#cjjl').show()"" name=""kscjjfjl"" value=""1"" "
			If ThumbnailsConfig(15) = "1" Then .Write (" checked")
			.Write ">启用&nbsp;&nbsp;"
	
			.Write "<input type=""radio"" onclick=""$('#cjjl').hide()"" name=""kscjjfjl"" value=""0"" "
			If ThumbnailsConfig(15) = "0" Then .Write (" checked")
			.Write ">"
			.Write "不启用<br/>"
			If ThumbnailsConfig(15) = "1" Then
			.Write "<div id='cjjl'>"
			Else
			.Write "<div id='cjjl' style='display:none'>"
			End If
		%>
		<script type="text/javascript">
    function doadd(num)
    {var i;
    var str="";
    var oldi=0;
    var j=0;
    oldi=parseInt(jQuery('#editnum').val());
    for(i=1;i<=num;i++)
    {
    j=parseInt(i)+oldi;
    str="<tr><td width=10% height=20> <div align=center><input type=hidden name=id value=0>"+j+"</div></td><td width=35%><input type=text name=cjitem size=10 class='txt' style='text-align:center'>%</td><td width=25%><input type=text name=cjscore style='text-align:center' class='txt' value=0 size=6> 分</td><td width='30%'><textarea class='txt' style='width:200px;height:80px' name='txtj' ></textarea></td></tr>";
     $("#cjjftr").append(str);
    }

     $('#editnum').val(j);
    }
	$(document).ready(function(){
	//doadd(1);
	});
	
    </script>
	<input type="button" name="Submit52" value="增加选项" class="button" onClick="javascript:doadd(1);"> 
	<table width="80%" border=0 cellspacing=1 cellpadding=3>
		   <tr  bgcolor="#E8F9FD" style="font-size:14px; font-weight:bold">
			 <td width='10%' height='30'  style="padding-left:35px;"> 编号</td>
			 <td width='35%'  style="padding-left:15px;">成绩(>=)</td>
			 <td style='25%'   style="padding-left:15px;">奖励积分</td>
             <td style='30%'   style="padding-left:15px;">鼓励语言</td>
		  </tr>
		  <%
		   Dim SNum:Snum=0
		   If Not KS.IsNul(ThumbnailsConfig(16)) Then
		     Dim SItemArr:SItemArr=split(ThumbnailsConfig(16),"§")
			 Snum=Ubound(SItemArr)+1
			 For I=0 To Snum-1
			   dim ssitem:ssitem=split(SItemArr(i)&"@","@")
			   response.write "<tr><td width=10% height=20> <div align=center><input type=hidden name=id value=0>" & (I+1) & "</div></td><td width=""35%""><input type=text name=cjitem size=10 class='txt' value=""" & ssitem(0) & """ style='text-align:center'>%</td><td width=""25%""><input type=text name=cjscore style='text-align:center' class='txt' value=""" & ssitem(1) & """ size=6> 分</td><td width=""30%""><textarea class='txt' style='width:200px;height:80px' name='txtj' >" & ssitem(2) & "</textarea></td></tr>"

			 Next
		   End If
		   %>
		  <tbody id="cjjftr">
		  
		  </tbody>
	</table>
		<input name="editnum" type="hidden" id="editnum" value="<%=Snum%>"> 

	<div class="tips">Tips:积分奖励的规则请按成绩从高到低的顺序设置,删除选项请将成绩项留空即可。</div>
		<%
		    .Write "</div>"
		.Write "</dd> "
		
		.Write "  <dd><div><strong>后台添加试卷启用word存图选项：</strong></div>"
		
		.Write "  <input type=""radio"" name=""FsoContentRule"" value=""1"" "
		If FsoContentRule = "1" Then .Write (" checked")
		.Write ">启用&nbsp;&nbsp;"

		.Write "<input type=""radio"" name=""FsoContentRule"" value=""0"" "
		If FsoContentRule = "0" Then .Write (" checked")
		.Write ">"
		.Write "不启用<span class='tips'>说明启用此功能，需要安装word存图插件(<a href='http://www.kesion.com/kfrz/13043.html' target='_blank'>查看</a>)，并且服务器需要支持asp.net 2.0环境。</span>"
		.Write "  </dd>"
		
		.Write "<dd><div>知识点启用选项：</div>"
		.Write "  <input type=""radio"" name=""SjzsdRule"" value=""1"" "
		If KS.ChkClng(ThumbnailsConfig(13))="1" Then .Write (" checked")
		.Write ">启用&nbsp;&nbsp;"
		.Write "<input type=""radio"" name=""SjzsdRule"" value=""0"" "
		If KS.ChkClng(ThumbnailsConfig(13)) = "0" Then .Write (" checked")
		.Write ">"
		.Write "不启用"
		.Write "    </dd>"
		
        .Write "<dd><div><strong>查看试题解释需要消费：</strong></div>"      
		.Write "	<label><input type='radio' name='jsxh'"
		if KS.ChkClng(ThumbnailsConfig(6))="0" then .write " checked"
		.Write " value='0'>不需要</label><br/>"
		.Write "<label><input type='radio' name='jsxh'"
		if KS.ChkClng(ThumbnailsConfig(6))="1" then .write " checked"
		.Write " value='1'>看一份试卷只需消费一次</label><br/>"
		.Write "<label><input type='radio' name='jsxh'"
		if KS.ChkClng(ThumbnailsConfig(6))="2" then .write " checked"
		.Write " value='2'>看一题消费一次</label>"
		.Write "<br/>查看试题解释每次消费<input class='txt' type='text' name='jsxhnum' size='4' style='text-align:center' value='" & ThumbnailsConfig(7) &"'/>元  间隔<input class='txt' type='text' name='rxhhour' size='4' style='text-align:center' value='" & ThumbnailsConfig(8) &"'/>小时后需要重复收复，不重复收费输入0"
		.Write "</dd>"
		.Write "<dd><div><strong>考试需要输验证码：</strong></div><label><input type='radio' name='ksyzm'"
		if KS.ChkClng(ThumbnailsConfig(11))="0" then .write " checked"
		.Write " value='0'>不需要</label>&nbsp;"
		.Write "<label><input type='radio' name='ksyzm'"
		if KS.ChkClng(ThumbnailsConfig(11))="1" then .write " checked"
		.Write " value='1'>需要</label>"
		
		.Write "</dd>"
		.Write "<dd><div><strong>考试首页分类从第：</strong></div><select name=""sjfl"">"
		 for i=1 to  KS.Chkclng(conn.execute("select max(tj) from ks_sjclass")(0))
		  if KS.ChkClng(ThumbnailsConfig(12))=i then
		  .write "<option value=" & i & " selected>第" & I & "层</option>"
		  else
		  .write "<option value=" & i & ">第" & I & "层</option>"
		  end if
		 next
		.Write "<select>生成。<span class='tips'> 指用标签{$GetClass}调用生成的分类，从第几层开始生成，建议选择第二层</span></dd>"
        .Write "</dl>"
		.Write "</div>"
		Else
		.Write " <div class=tab-page id=fso-page>"
		.Write "  <H2 class=tab>生成选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""fso-page"" ) );"
		.Write "	</SCRIPT>"

		.Write "<dl class=""dtable"">"
		.Write "<dd><div>本模型运行模式：</div>"
		.Write "<strong>PC版：</strong><br/>"
		.Write "  <input type=""radio"" name=""FsoHtmlTF"" value=""0"" "
		If FsoHtmlTF = 0 Then .Write (" checked")
		.Write ">动态asp<br/>"

		.Write "<input type=""radio"" name=""FsoHtmlTF"" value=""1"" "
		If FsoHtmlTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "栏目页及内容页都生成HTML<br/>"
		
		.Write "<input type=""radio"" name=""FsoHtmlTF"" value=""2"" "
		If FsoHtmlTF = 2 Then .Write (" checked")
		.Write ">"
		.Write "栏目页不生成,内容页生成HTML(<font color=red>推荐</font>)"
		
		If BasicType<> 4 and KS.WSetting(0)="1" Then
		  .Write "<font>"
		Else
		  .Write "<font style='display:none'>"
		End If
		.Write "<br/><strong>手机版：</strong><br/>"

        .Write "  <input type=""radio"" name=""MFsoHtmlTF"" value=""0"" "
		If KS.ChkClng(ThumbnailsConfig(28)) = 0  Then .Write (" checked")
		.Write ">动态asp<br/>"

		.Write "<input type=""radio"" name=""MFsoHtmlTF"" value=""1"" "
		If KS.ChkClng(ThumbnailsConfig(28)) = 1 Then .Write (" checked")
		.Write ">"
		.Write "内容页生成HTML"
		.Write " </font>   </dd>"
		
	
		
		
		.Write "<font id='staticarea'>"
		.Write "    <dd><div>伪静态设置：</div>"
		.Write "  <input type=""radio"" name=""StaticTF"" value=""0"" "
		If StaticTF = 0 Then .Write (" checked")
		.Write ">"
		.Write "不启用"
		.Write "  <input type=""radio"" name=""StaticTF"" value=""1"" "
		If StaticTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "伪静态(带问号,不需要装组件)"
		.Write "  <input type=""radio"" name=""StaticTF"" value=""2"" "
		If StaticTF = 2 Then .Write (" checked")
		.Write ">"
		.Write "伪静态(需要装ISAPI_Rewrite组件)"
        .Write "<span>这里需要设置不生成静态才有效,建议流量大的网站直接启用全部生成静态,而不是使用伪静态</span>"
		.Write "  </dd>"
		.Write " </font>"
		.Write "<font id='fsoarea'>"
		.Write "    <dd><div>添加文档，同时发布HTML选项：</div>"
		.Write "     <input type=""radio"" name=""RefreshFlag"" value=""1"" "
		If RefreshFlag = 1 Then .Write (" checked")
		.Write ">"
		.Write "仅发布内容页 <br>"
		.Write "          <input type=""radio"" name=""RefreshFlag"" value=""2"" "
		If RefreshFlag = 2 Then .Write (" checked")
		.Write ">发布栏目页+内容页<font color=red>(建议)</font><br>"		
		.Write "          <input type=""radio"" name=""RefreshFlag"" value=""3"" "
		If RefreshFlag = 3 Then .Write (" checked")
		.Write ">发布首页+栏目页+内容页"
		.Write "    </dd>"	
		.Write "    <dd><div>自动生成列表分页数：</div><input class='txt textbox' style='text-align:center;width:60px' type='text' value='" & FsoListNum & "' name='FsoListNum' size='10' /><span>这里设置生成栏目列表分页时自动生成的分页数，如果你的网站数据量较大，建议输入一个较小的数字，小数据量的网站可以不用限制，直接设置为0</span></dd>"	
	    .Write "<dd><div>生成的总目录(PC版)：</div>"      
		.Write "<input class='txt textbox' name='FsoFolder' type='text' value='" & FsoFolder & "' size='30'/><span class='tips'>*用于生成静态html存放的目录，只能以字母和数字的组合,必须以""/""结束</span>"
		.Write "</dd>"
		.Write "<dd><div>生成的栏目页规则(PC版)：</div>"   
		.Write "<input type=""radio"" name=""FsoClassListRule"" value=""1"" "
		If FsoClassListRule = 1 Then .Write (" checked")
		.Write ">按目录级别顺序结构生成列表页<br>"
		.Write " &nbsp;<font color=blue>如：第1页为/article/aaa/bbb/ccc/index.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/article/aaa/bbb/ccc/index_2.html</font>"
		.Write "  <br><input type=""radio"" name=""FsoClassListRule"" value=""2"" "
		If FsoClassListRule = 2 Then .Write (" checked")
		.Write ">所有栏目页都生成在模型总生成目录下面<font color=red>(有利于SEO)</font><br>"
		.Write " &nbsp;<font color=green>如栏目ID号为100则生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/list_100.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/list_100_2.html</font>"
		.Write " <br>&nbsp;<font color=green>如栏目ID号为101则生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/list_101.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/list_101_2.html</font>"
		.Write "  <br><input type=""radio"" name=""FsoClassListRule"" value=""3"" "
		If FsoClassListRule = 3 Then .Write (" checked")
		.Write ">本模型下的一级栏目生成在本频道下的Index.html,子栏目按如下规则生成<br>"
		.Write " &nbsp;<font color=green>如一级栏目 ""教育频道"" 英文名称：""edu"",那么生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/edu/index.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/edu/index_2.html</font>"
		.Write " <br>&nbsp;<font color=green>二级及以上的栏目(即""教育频道"")下的栏目,如栏目ID号为101则生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/edu/list_101.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/edu/list_101_2.html</font>"
		.Write "  <br><input type=""radio"" name=""FsoClassListRule"" value=""4"" "
		If FsoClassListRule = 4 Then .Write (" checked")
		.Write ">所有栏目页都生成在模型总生成目录下面<font color=red>(新增）</font><br>"
		.Write " <font color=blue>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/自定义列表前缀.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/自定义列表前缀_2.html</font>"
		
		.Write "</dd>"

		.Write "<dd><div><strong>生成的列表页的前缀字符(PC版)：</strong></div><input class='txt textbox' name='FsoClassPreTag' type='text' value='" & FsoClassPreTag & "' size='30'> <span class='tips'>*如list,show等</span><br/><span class='tips'>可用标签：<br/>{$ClassEname}-本栏目英文名<br/>{$ClassID}-本栏目小ID<br/> {$BigClassID}-本栏目大ID<br/>{$TopClassEname}-一级栏目英文名 "
		.Write "</dd>"


		.Write "<dd><div>生成的内容页目录规则(PC版)：</div>"      
		.Write " <input class='txt textbox' name='FsoContentRule' type='text' value='" & FsoContentRule & "' size='30'>&nbsp;"
		.Write "     <select name='srule' onchange='if (this.value!=""""){ $(""input[name=FsoContentRule]"").val(this.value);}'><option value=''>------快速选择内容页生成结构------</option>"
		.Write "     <option value='View_'>View_</option>"
		.Write "     <option value='{$ClassDir}'>{$ClassDir}</option>"
		.Write "     <option value='{$ChannelEname}/{$ClassEname}_{$ClassID}_'>{$ChannelEname}/{$ClassEname}_{$ClassID}_</option>"
		.Write "     <option value='{$ClassEname}_{$ClassID}_'>{$ClassEname}_{$ClassID}_(推荐)</option>"
		.Write "     <option value='{$ClassDir}{$ClassEname}_{$ClassID}_'>{$ClassDir}{$ClassEname}_{$ClassID}_</option>"
		.Write "     <option value='{$Year}_{$Month}_{$Day}/'>{$Year}_{$Month}_{$Day}/</option>"
		.Write "     <option value='{$Year}_{$Month}/'>{$Year}_{$Month}/</option>"
		.Write "     <option value='{$Year}/{$Month}/{$Day}/'>{$Year}_{$Month}_{$Day}/</option>"
		.Write "     <option value='{$Year}/{$Month}/'>{$Year}_{$Month}/</option>"
		.Write "   </select><br><font color=red>可选项（允许留空）</font><br><span class='tips'>可用标签：一级频道名称{$ChannelEname},栏目路径{$ClassDir} 栏目ID号{$ClassID} 栏目英文名称{$ClassEname} 文档ID号{$InfoID}  文档添加年份{$Year} 文档添加月份{$Month} 文档添加日{$Day}</span><br> "
		.Write " </dd>"
		.Write "</font>"
        .Write "</dl>"
		.Write "</div>"
	End If
		.Write " <div class=tab-page id=upfile-page>"
		.Write "  <H2 class=tab>上传选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""upfile-page"" ) );"
		.Write "	</SCRIPT>"

		.Write "<dl class=""dtable""><dd><div>管理员是否允许上传文件：</div><input type=""radio"" name=""UpFilesTF"" value=""1"" "
		If UpFilesTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "允许"
		.Write "  <input type=""radio"" name=""UpFilesTF"" value=""0"""
		If UpFilesTF = 0 Then .Write (" checked")
		.Write ">"
		.Write "  不允许"
		.Write "    </dd>"
		.Write "   <dd><div>缩略图选项：</div>" & vbCrLf
					    
		.Write "黄金分割点：<input class='txt textbox' type='text' value='" & ThumbnailsConfig(0) & "' name='GoldenPoint' size='4' style='width:50px;text-align:center'> 宽度：<input class='txt textbox' type='text' value='" & ThumbnailsConfig(1) & "' name='ThumbsWidth' size='8' style='width:50px;text-align:center'>px 高度：<input class='txt textbox' type='text' value='" & ThumbnailsConfig(2) & "' name='ThumbsHeight' size='8' style='width:50px;text-align:center'>px"
		.Write "            <span>当基本信息设置开启自动生成缩略图功能时才有效</span> <br/><font color=red>tips:如果高度设置为0,则生成的高度将由您设置的宽度自动约束决定(类似photoshop软件的自动约束)</font>>"
		.Write "    </dd>" & vbCrLf
		.Write "    <dd style='display:none'><div>后台文件上传目录：</div><input name=""UpfilesDir"" class='txt' type=""text"" id=""UpfilesDir"" value=""" & UpfilesDir & """ size=""30""></dd>"
		
		.Write "    <dd><div>会员中心上传设置：</div> <input type=""radio"" name=""UserUpFilesTF"" value=""0"""
		If UserUpFilesTF = 0 Then .Write (" checked")
		.Write ">关闭上传<br/>"
		.Write "<input type=""radio"" name=""UserUpFilesTF"" value=""1"" "
		If UserUpFilesTF = 1 Then .Write (" checked")
		.Write ">只允许会员上传<br/>"
		.Write "<input type=""radio"" name=""UserUpFilesTF"" value=""2"" "
		If UserUpFilesTF = 2 Then .Write (" checked")
		.Write ">允许所有人上传，包括游客(匿名投稿)"
		
		.Write "  </dd>"
		.Write "    <dd style='display:none'>"
		.Write "     <div><strong>会员文件上传目录：</strong></div><input class='txt' name=""UserUpfilesDir"" type=""text"" id=""UserUpfilesDir"" value=""" & UserUpfilesDir & """ size=""30""><br><b>提示：</b><br><font color=red>1、会员目录构成规则：系统设置的总上传目录/User/会员名称;<br>2、上传目录必须以/结束;</font>>"
		.Write "    </dd>"
		
		.Write "    <dd><div>允许会员选择上传文件：</div><input type=""radio"" name=""UserSelectFilesTF"" value=""1"" "
		If UserSelectFilesTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "允许"
		.Write "  <input type=""radio"" name=""UserSelectFilesTF"" value=""0"""
		If UserSelectFilesTF = 0 Then .Write (" checked")
		.Write ">"
		.Write "  不允许</dd>"
		
		.Write "    <dd><div>允许上传的最大文件大小：</div><input name=""UpfilesSize"" class='txt textbox' onBlur=""CheckNumber(this,'允许上传最大文件大小');"" type=""text"" id=""UpfilesSize"" value=""" & UpfilesSize & """ size=""10"" style=""width:80px"">"
		.Write "      KB 　 <span>提示：1 KB = 1024 Byte，1 MB = 1024 KB</span>"
		.Write "    </dd>"
		.Write "    <dd><div>允许上传的文件类型：<font color='#ff0000' style='font-size:12px;font-weight:normal'>多种文件类型之间以""|""分隔</font>"
		.Write "          </div><table width=""98%"" border=""0"">"
		.Write "        <tr>"
		.Write "          <td width=""80"" height=""25"" align=""right"">图片类型:</td>"
		.Write "          <td><input class='txt textbox' name=""AllowUpPhotoType"" type=""text"" id=""AllowUpPhotoType"" value=""" & AllowUpPhotoType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">Flash 文件:</td>"
		.Write "          <td><input class='txt textbox' name=""AllowUpFlashType"" type=""text"" id=""AllowUpFlashType"" value=""" & AllowUpFlashType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">Windows 媒体: </td>"
		.Write "          <td><input class='txt textbox'  name=""AllowUpMediaType"" type=""text"" id=""AllowUpMediaType"" value=""" & AllowUpMediaType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">Real 媒体: </td>"
		.Write "          <td><input class='txt textbox' name=""AllowUpRealType"" type=""text"" id=""AllowUpRealType"" value=""" & AllowUpRealType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">其它文件:</td>"
		.Write "          <td><input class='txt textbox' name=""AllowUpOtherType"" type=""text"" id=""AllowUpOtherType"" value=""" & AllowUpOtherType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "      </table>"
		.Write "    </dd>"
        .Write "</dl>"
		.Write "</div>"

		.Write " <div class=tab-page id=tougao-page>"
		.Write "  <H2 class=tab>投稿选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "	  tabPane1.addTabPage( document.getElementById( ""tougao-page"" ) );"
		.Write "	</SCRIPT>"
		.Write "<dl class=""dtable"">"
		.Write "      <dd><div>前台会员投稿总开关：</div>"
		.Write "          <input type=""radio"" name=""UserTF"" value=""0"" "
		If UserTF = 0 Then .Write (" checked")
		.Write ">关闭会员投稿 <br>"
		.Write " <input type=""radio"" name=""UserTF"" value=""1"" "
		If UserTF = 1 Then .Write (" checked")
		.Write ">只允许注册会员可以投稿,具体可投稿的栏目依栏目设置而定<br/>"
		.Write " <input type=""radio"" name=""UserTF"" value=""2"" "
		If UserTF = 2 Then .Write (" checked")
		.Write ">允许所有用户投稿（包括游客）,具体可投稿的栏目依栏目设置而定<br>"
		
		.Write " <input type=""radio"" onclick=""$('#sGroup').show();"" name=""UserTF"" value=""3"" "
		If UserTF = 3 Then .Write (" checked")
		.Write ">按会员组投稿 （指定会员组可以投稿）<br>"
		
		
		 .Write "<table border='0' align=center width='90%'>"
		 .Write " <tr>"
		 IF UserTF = 3 Then
			.Write " <td id='sGroup'>"
		 Else
			 .Write "<td id='sGroup' style='display:none'>"
		 End If
		 .Write KS.GetUserGroup_CheckBox("GroupID",ThumbnailsConfig(17),5)
		 .Write " </td>"
		 .Write "  </tr></table>"
		.Write "    </dd>"
		.Write "    <dd><div>会员中心投稿菜单显示名称：</div>"
		.Write "    <input class='txt textbox' name=""ModelShortName"" type=""text"" id=""ModelShortName"" value=""" & ModelShortName & """ size=""50""> <span class='tips'>如：文章，新闻，房源等，建议取两个汉字名称</span>"
		.Write "    </dd>"
		 Dim CurrPath:CurrPath = KS.GetUpFilesDir():If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath, Len(CurrPath) - 1)
		.Write "    <dd><div>会员中心投稿菜单图标地址：</div>"
		.Write "    <input class='txt textbox' name=""ModelIco"" type=""text"" id=""ModelIco"" value=""" & ModelIco & """ size=""50""> <input class=""button""  type='button' name='Submit' value='选择图片...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & CurrPath & "',550,290,window,$('#ModelIco')[0]);""> <span class='tips'>如：/user/images/ico1.gif</span>"
		.Write "    </dd>"
		
		.Write "    <dd><div>新注册会员：</div>"
		.Write "    <input class='txt textbox' name=""PubTimeLimit"" type=""text"" style=""width:60px;text-align:center"" id=""PubTimeLimit"" value=""" & PubTimeLimit & """ size=""6"">分钟后才可以在此模型投稿"
		.Write "    </dd>"
		
		.Write "    <dd><div>会员投稿增加：</div> 资金<input class='txt textbox' style='width:60px;text-align:center' name=""UserAddMoney"" type=""text"" id=""UserAddMoney"" value=""" & UserAddMoney & """ size=""6"">元  点券<input class='txt textbox' style='width:60px;text-align:center' name=""UserAddPoint"" type=""text"" id=""UserAddPoint"" value=""" & UserAddPoint & """ size=""6"">点  积分<input class='txt textbox'  name=""UserAddScore"" type=""text"" id=""UserAddScore"" value=""" & UserAddScore & """ style='width:60px;text-align:center' size=""6"">分<span>为0时不增加,可设置成负数,表示投稿要消费</span>"
		.Write "    </dd>"
		
		.Write "    <dd><div>允许会员刷新添加时间：</div>"
		.Write " <input type=""radio"" name=""RefreshTimeTF"" value=""0"" "
		If ThumbnailsConfig(3) = "0" Then .Write (" checked")
		.Write ">不允许 "
		.Write "          <input type=""radio"" name=""RefreshTimeTF"" value=""1"" "
		If ThumbnailsConfig(3) = "1" Then .Write (" checked")
		.Write ">允许</dd>"
		

		.Write "    <dd><div>审核过的稿件是否允许修改：</div>"
		.Write " <input type=""radio"" name=""UserEditTF"" value=""0"" "
		If UserEditTF = 0 Then .Write (" checked")
		.Write ">不允许<font color=red>(建议)</font><br>"
		.Write "          <input type=""radio"" name=""UserEditTF"" value=""1"" "
		If UserEditTF = 1 Then .Write (" checked")
		.Write ">允许，但修改后自动转为未审(<font color=red>如果投稿要增加积分等,会导致重复收费</font>)<br>"
		.Write "          <input type=""radio"" name=""UserEditTF"" value=""2"" "
		If UserEditTF =2 Then .Write (" checked")
		.Write ">允许，修改后仍为已审状态（<font color=red>不推荐,如果投稿要增加积分等,会导致重复收费</font>）"
        .Write "      </dd>"
		.Write "    <dd><div>投稿栏目显示方式：</div>"
		.Write " <input type=""radio"" name=""UserClassStyle"" value=""0"" "
		If UserClassStyle = 0 Then .Write (" checked")
		.Write ">仅显示有权限的栏目（下拉方式）<br>"
		.Write "          <input type=""radio"" name=""UserClassStyle"" value=""3"" "
		If UserClassStyle = 3 Then .Write (" checked")
		.Write ">多级联动下拉（适合于只有二至三级栏目结构的模型）"
        .Write "      </dd>"
		
		
		.Write "    <dd><div>会员中心发布的信息是否需要审核：</div>"
		.Write " <input type=""radio"" name=""InfoVerificTF"" value=""0"" "
		If InfoVerificTF = 0 Then .Write (" checked")
		.Write ">需要后台人工审核<br>"
		.Write "          <input type=""radio"" name=""InfoVerificTF"" value=""1"" "
		If InfoVerificTF = 1 Then .Write (" checked")
		.Write ">不需要审核（但不直接生成内容页HTML）<br>"
		.Write "          <input type=""radio"" name=""InfoVerificTF"" value=""2"" "
		If InfoVerificTF = 2 Then .Write (" checked")
		.Write ">不需要审核（当有启用生成静态HTML，直接生成内容页） </dd>"
		
		If BasicType=1 Or BasicType=2 Or BasicType=3 Then
		.Write "    <dd><div>自动生成投稿录入表单:</div>"
		.Write "    <label><input type='checkbox' name='autocreate' id='autocreate' value='1' onClick=""LoadTemplate(this.checked)"">自动生成</label> <span>提示：第一次生成模板，可以点此自动生成！</span>"
		%>
        <div class="clear"></div>
		<script language = 'JavaScript'>
			function LoadTemplate(v){   
					   if (v==true)
					    { 
							$.ajax({
								  url: 'KS.Model.asp',
								  cache: false,
								  data: "action=createtemplate&channelid=<%=ChannelID%>",
								  success: function(s){
									  $('#Content').val(s);
								  }
								});
							 return; 
						}
						else
						{
						  $('#Content').val('');
						}
			}	

		    function show_ln(txt_ln,txt_main){
			            var txt_ln  = document.getElementById(txt_ln);
			            var txt_main  = document.getElementById(txt_main);
			            txt_ln.scrollTop = txt_main.scrollTop;
			            while(txt_ln.scrollTop != txt_main.scrollTop)
			            {
				            txt_ln.value += (i++) + '\n';
				            txt_ln.scrollTop = txt_main.scrollTop;
			            }
			            return;
		            }
		  function editTab(){
			            
			            }
		            //-->
		 </script>

			
			 <textarea id='txt_ln' name='rollContent' cols='6' style='overflow:hidden;height:280px;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;' readonly><%
			Dim XmlForm:XmlForm=LFCls.GetConfigFromXML("modelinputform","/inputform/model",ChannelID)
			If KS.IsNul(XmlForm) Then XmlForm=""
			 
		 Dim N
		 For N=1 To 3000
			Response.Write N & "&#13;&#10;"
		 Next
		 On Error Resume Next
		 %>
		 </textarea>
		 <textarea name='Content' id="Content" style="width:700px;height:285px;" ROWS='15' onkeydown='editTab()' onscroll="show_ln('txt_ln','Content')" wrap='on'><%=Server.HTMLEncode(XmlForm)%></textarea>
         	<div class="clear"></div>
            <span>不想自定义可以留空,否则添加/变更字段需要重新生成表单模板</span>
			 </dd>
		<%
		End If

        .Write "</dl>"
		.Write "</div>"
		
		If ChannelID<>9 Then
		.Write " <div class=tab-page id=digg-page>"
		If KS.GetAppStatus("digmood")=TRUE THEN
		.Write "  <H2 class=tab>Digg选项</H2>"
		Else
		.Write "  <H2 class=tab style=""display:none"">Digg选项</H2>"
		End If
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""digg-page"" ) );"
		.Write "	</SCRIPT>"
		.Write "<dl class='dtable'>"
		.Write "      <dd><div>是否允许游客DIGG：</div>"
		.Write " <input type=""radio"" name=""DiggByVisitor"" value=""1"" "
		If DiggByVisitor = 1 Then .Write (" checked")
		.Write ">允许"
		.Write "          <input type=""radio"" name=""DiggByVisitor"" value=""0"" "
		If DiggByVisitor = 0 Then .Write (" checked")
		.Write ">不允许"
		.Write "    </dd>"
		.Write "    <dd><div>是否启用单IP限制：</div>"
		.Write " <input type=""radio"" name=""DiggByIP"" value=""1"" "
		If DiggByIP = 1 Then .Write (" checked")
		.Write ">启用"
		.Write "          <input type=""radio"" name=""DiggByIP"" value=""0"" "
		If DiggByIP = 0 Then .Write (" checked")
		.Write ">不启用"
		.Write "      <span>若启用单IP限制，每个IP的用户只能对每个项目Digg一次</span>"
		.Write "    </dd>"
		.Write "    <dd><div>会员是否允许重复DIGG：</div>"
		.Write " <input type=""radio"" name=""DiggRepeat"" value=""1"" "
		If DiggRepeat = 1 Then .Write (" checked")
		.Write ">允许"
		.Write "          <input type=""radio"" name=""DiggRepeat"" value=""0"" "
		If DiggRepeat = 0 Then .Write (" checked")
		.Write ">不允许"
		.Write "      <span>启用IP限制时，始终不允许</span>"
		.Write "    </dd>"
		.Write "    <dd><div>次数选项：</div>"
		.Write "         每DIGG一下自动增加<input type=""text"" class='txt textbox' size=""6"" style=""width:60px;text-align:center""  name=""DiggPerTimes"" value=""" & DiggPerTimes & """>次 "
		.Write "      </dd>"
        .Write "</dl>"
        .Write "</div>"	
	End If	 
		 
		.Write " <div class=tab-page id=detail-page>"
		.Write "  <H2 class=tab>其它参数</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""detail-page"" ) );"
		.Write "	</SCRIPT>"

		.Write "<input type=""hidden"" value=""Edit"" name=""Flag"">"

		.Write "    <dl class=""dtable""><dd><div>本模型计费方式：</div><input type=""radio"" name=""ChargeType"" value=""0"" "
		If ChargeType = 0 Then .Write (" checked")
		.Write ">"
		.Write "       " & KS.Setting(45)
		.Write "          <input type=""radio"" name=""ChargeType"" value=""1"" "
		If ChargeType = 1 Then .Write (" checked")
		.Write ">"
		.Write "        资金(人民币)"		
		.Write "          <input type=""radio"" name=""ChargeType"" value=""2"" "
		If ChargeType = 2 Then .Write (" checked")
		.Write ">"
		.Write "        积分<span>如文章/图片/下载等设置需要消费才可以查看,将以这里设置的计费标准扣费,一旦设置建议不要修改,此次设置对商城模型无效</span>"
		.Write "    </dd>"	
		.Write "    <dd" 
		if channelid=9 then response.write " style='display:none'>" else response.write ">"
		.Write "     <div>下载本模型附件费用：</div><input class='txt textbox' type=""text"" style=""width:60px;text-align:center"" size=8 name=""AnnexPoint"" value=""" & AnnexPoint & """> 24小时内下载不重复扣费,不限制请输入0"
		.Write "    </dd>"
			
	
		.Write "    <dd" 
		if channelid=9 then response.write " style='display:none'>" else response.write ">"
		.Write "     <div>最新信息标志：</div>"
		.Write "     <input class='txt textbox' type=""text"" style=""width:60px;text-align:center"" size=8 name=""LatestNewDay"" value=""" & LatestNewDay & """>天内添加的信息标志为最新信息"
		.Write "    </dd>"
		.Write "    <dd><div>生成的文档内容里的图片是否格式化：</div>"
        .Write "<input type='radio' name='formatcontentimg' value='0'"
		
		 if KS.ChkClng(ThumbnailsConfig(27))="0" then .Write " checked"
		.Write ">不启用"
		.Write "<input type='radio' name='formatcontentimg' value='1'"
		if ThumbnailsConfig(27)="1" then .Write " checked"
		.Write ">启用"
		
		.Write " <span>TIPS:启用格式化后，前台生成的文章内容里的图片会随鼠标缩放。</span>   </dd>"
		
		
		.Write "    <dd><div>后台管理设置：</div>后台每页显示：</strong> <input class='txt textbox' type=""text"" style='width:60px;text-align:center' size='4' name=""MaxPerPage"" value=""" & MaxPerPage & """>条信息"
		.Write "<br/>显示按栏目筛选："
		.Write "<input type='radio' name='classsearch' value='0'"
		 if ThumbnailsConfig(24)="0" then .Write " checked"
		.Write ">不显示"
		.Write "<input type='radio' name='classsearch' value='1'"
		if ThumbnailsConfig(24)="1" then .Write " checked"
		.Write ">显示但不带文档数量"
		.Write "<input type='radio' name='classsearch' value='2'"
		if ThumbnailsConfig(24)="2" then .Write " checked"
		.Write ">显示带文档数量(数据量大时，性能有所影响）"
		
		.Write "<br/>审核模式："
		.Write "<input type='radio' name='VerifyJB' value='0'"
		 if ThumbnailsConfig(25)="0" then .Write " checked"
		.Write ">一级审核"
		.Write "<input type='radio' name='VerifyJB' value='1'"
		if ThumbnailsConfig(25)="1" then .Write " checked"
		.Write ">二级审核"
		.Write "<br/>后台添加文档布局："
		.Write "<input type='radio' name='AddDocStyle' value='0'"
		 if ThumbnailsConfig(30)="0" then .Write " checked"
		.Write ">上下布局"
		.Write "<input type='radio' name='AddDocStyle' value='1'"
		if ThumbnailsConfig(30)="1" then .Write " checked"
		.Write ">左右布局"

		
		.Write "   </dd>"
		
		.Write "    <dd" 
		if channelid=9 then response.write " style='display:none'>" else response.write ">"
		.Write "      <div>本模型启用回收站功能：</div>"
		.Write " <input type=""radio"" name=""DelTF"" value=""1"" "
		If ThumbnailsConfig(5) = "1" Then .Write (" checked")
		.Write ">不启用 "
		.Write "          <input type=""radio"" name=""DelTF"" value=""0"" "
		If ThumbnailsConfig(5) = "0" Then .Write (" checked")
		.Write ">启用<span>启用回收站后，则删除文档将放入回收站里，可以在回收站中还原。</span>"
		.Write "    </dd>"
		
		.Write "    <dd>"
		.Write "      <div>本模型前台允许搜索：</div>"
		.Write " <input type=""radio"" name=""SearchTF"" value=""1"" "
		If ThumbnailsConfig(29) = "1" Then .Write (" checked")
		.Write ">允许 "
		.Write "          <input type=""radio"" name=""SearchTF"" value=""0"" "
		If ThumbnailsConfig(29) = "0" Then .Write (" checked")
		.Write ">不允许<span>设置允许搜索，才能才能使用筛选功能。</span>"
		.Write "    </dd>"
		
		
		.Write "    <dd" 
		if channelid=9 then response.write " style='display:none'>" else response.write ">"
		.Write "     <div>评论设置：</div>"
		.Write "     <table width=""98%"" border=""0"">"
		.Write "     <tr valign=""middle"">"
		.Write "      <td width=""150"" height=""30"" width='160' align=""right"">本模型评论系统设置：</td>"
		.Write "      <td height=""30"">"
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""0"" "
		If VerificCommentTF = 0 Then .Write (" checked")
		.Write ">关闭本模型的所有信息评论<br>"

		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""1"" "
		If VerificCommentTF = 1 Then .Write (" checked")
		.Write ">本模型只允许会员评论，且评论内容需要后台的审核<br>"
		
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""2"" "
		If VerificCommentTF = 2 Then .Write (" checked")
		.Write ">本模型只允许会员评论，且评论内容不需要后台审核<br>"
		
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""3"" "
		If VerificCommentTF = 3 Then .Write (" checked")
		.Write ">本模型允许会员，游客评论，且评论内容需要后台审核<br>"
		
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""4"" "
		If VerificCommentTF = 4 Then .Write (" checked")
		.Write ">本模型允许会员，游客评论，且评论内容不需要后台审核<br/>"
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""5"" "
		If VerificCommentTF = 5 Then .Write (" checked")
		.Write ">本模型开放所有用户评论，但会员评论不需审核，游客需要审核"

		

		.Write "             </td>"
		.Write "    </tr>"		
		.Write "    <tr valign=""middle"">"
		.Write "      <td height=""30"" style=""text-align:right"">评论需要验证码：</td>"
		.Write "      <td height=""30""> <input type=""radio"" name=""CommentVF"" value=""1"" "
		If CommentVF = 1 Then .Write (" checked")
		.Write ">"
		.Write "        是"
		.Write "          <input type=""radio"" name=""CommentVF"" value=""0"" "
		If CommentVF = 0 Then .Write (" checked")
		.Write ">"
		.Write "          否        </td>"
		.Write "    </tr>"
		.Write "    <tr valign=""middle"">"
		.Write "      <td height=""30"" align=""right"">评论字数控制：</td>"
		.Write "      <td height=""30""> <input style=""width:60px;text-align:center"" class='txt textbox' name=""CommentLen"" type=""text"" value=""" & CommentLen & """ size=""6"">个字,同一个IP<input class='txt textbox' name=""CommentTime"" type=""text"" style=""width:60px;text-align:center"" value=""" & ThumbnailsConfig(4) & """ size=""6"">小时内对同一篇文档只能发表一次。不限制请输入""0""</td>"
		.Write "    </tr>"		
		.Write "    <tr valign=""middle"">"
		.Write "      <td height=""30"" align=""right"">内容页每页显示评论条数：</td>"
		.Write "      <td height=""30""> <input style=""width:60px;text-align:center"" class='txt textbox' name=""CommentPerPage1"" type=""text"" value=""" & ThumbnailsConfig(22) & """ size=""6"">条,更多评论页面每页显示<input class='txt textbox' name=""CommentPerPage2"" type=""text"" style=""width:60px;text-align:center"" value=""" & ThumbnailsConfig(23) & """ size=""6"">条</td>"
		.Write "    </tr>"		
		.Write "    <tr valign=""middle"">"
		.Write "      <td height=""30"" align=""right"">评论页模板：</td>"
		.Write "      <td height=""30""><input class='txt textbox' name=""CommentTemplate"" style=""width:300px"" id=""CommentTemplate"" type=""text"" value=""" & CommentTemplate & """>&nbsp;" & KSCls.Get_KS_T_C("$('#CommentTemplate')[0]") & "</td>"
		.Write "    </tr>"			
		.Write "      </table>"
		.Write "    </dd>"	
		If KS.WSetting(0)="1" Then
		.Write "    <dd style='display:none'>"
		.Write "      <div>3G搜索页模板：</div><input class='txt textbox' style=""width:300px""  name=""WapSearchTemplate"" id=""WapSearchTemplate"" type=""text"" value=""" & WapSearchTemplate & """>&nbsp;" & KSCls.Get_KS_T_C("$('#WapSearchTemplate')[0]")
		.Write "    </dd>"
		End If
		.Write "  </dl>"
		.Write "</div>"
		.Write "</form>"
		End With
		End Sub
		
		Sub ChannelSave()
		     On Error Resume Next
		    Dim ModelEname,ThumbnailsConfig,ChannelTable,I,OpName,ItemName,ChannelID:ChannelID=KS.ChkClng(KS.G("ChannelID"))
            If KS.IsNul(KS.G("ChannelName")) Then
				   Call KS.AlertHistory("请输入模型名称!",-1)
				   Exit Sub
			End If
            If KS.IsNul(KS.G("ModelEName")) And OpName="添加" Then
				   Call KS.AlertHistory("请输入模型英文名称!",-1)
				   Exit Sub
			End If
			ItemName=KS.G("ItemName")
			
			'============考试成绩奖励===========================================
			Dim CJItem,cjscore,CJItemStr,txtj
			if Not KS.IsNul(KS.FilterIds(request("cjitem"))) Then
			  CJItem=Split(KS.FilterIds(request("cjitem")),",")
			  cjscore=Split(Trim(request("cjscore")),",")
			  txtj=Split(Trim(request("txtj"))&"",",")		
			  For I=0 To Ubound(CJItem)
			    If Not KS.IsNul(CJItem(i)) Then
					If CJItemStr="" Then
					 CJItemStr=KS.ChkClng(CJItem(i))&"@" & KS.ChkClng(cjscore(i))&"@"&Trim(txtj(i))
					Else
					 CJItemStr=CJItemStr&"§" & KS.ChkClng(CJItem(i))&"@" & KS.ChkClng(cjscore(i))&"@"&Trim(txtj(i))
					End If
				End If
			  Next
			End If
			
			'=====================================================================
			dim imgFsoFolder1,imgFsoFolder2
			imgFsoFolder1=KS.G("imgFsoFolder1"):imgFsoFolder2=KS.G("imgFsoFolder2")
			ThumbnailsConfig=Request.Form("GoldenPoint") & "|" & KS.ChkClng(Request.Form("ThumbsWidth")) & "|" & KS.ChkClng(Request.Form("ThumbsHeight")) & "|" & KS.ChkClng(Request.Form("RefreshTimeTF"))& "|" & KS.ChkClng(Request.Form("CommentTime")) & "|" & KS.ChkClng(Request.Form("DelTF"))& "|" & KS.ChkClng(Request("jsxh")) & "|" & Request("jsxhnum") & "|" & KS.ChkClng(Request("rxhhour")) & "|" & request("rightyy") & "|" & request("wrongyy")& "|" & KS.ChkClng(request("ksyzm")) & "|" & ks.chkclng(request("sjfl")) & "|" & ks.chkclng(request("SJzsdRule")) & "|" & KS.ChkClng(request("rclxjfjl"))& "|" & KS.ChkClng(request("kscjjfjl")) & "|" & CJItemStr & "|" & Request.Form("GroupID")  & "|" & ks.chkclng(request("ksxd"))   & "|" & ks.chkclng(request("ksxdsh")) & "|" & imgFsoFolder1& "|" & imgFsoFolder2&"|" & KS.ChkClng(Request("CommentPerPage1")) &"|" & KS.ChkClng(Request("CommentPerPage2"))
			ThumbnailsConfig=ThumbnailsConfig & "|"& KS.ChkCLng(Request("classsearch")) '24
			ThumbnailsConfig=ThumbnailsConfig & "|"& KS.ChkCLng(Request("verifyJB")) '25
			ThumbnailsConfig=ThumbnailsConfig & "|"&Request("FolderName") '26
			ThumbnailsConfig=ThumbnailsConfig & "|"&Request("FormatContentImg") '27
			ThumbnailsConfig=ThumbnailsConfig & "|"& KS.ChkClng(Request("MFsoHtmlTF")) '28
			ThumbnailsConfig=ThumbnailsConfig & "|"& KS.ChkClng(Request("SearchTF")) '29
			ThumbnailsConfig=ThumbnailsConfig & "|"& KS.ChkClng(Request("AddDocStyle")) '30 后台添加文档排序方式
			Dim MaxOrderID:MaxOrderID=0
			If ChannelID=0 Then
			   MaxOrderID=KS.ChkClng(Conn.Execute("Select max(orderid) From KS_Channel")(0))+1
			End If
			
		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_Channel Where ChannelID=" & ChannelID,Conn,1,3
			If  RS.Eof And RS.Bof Then
			    RS.AddNew
				 OpName       = "添加"
				 ChannelTable = "KS_U_" & KS.G("ChannelTable")
				 ModelEname   = Replace(Replace(Replace(KS.G("ModelEname"), "\","/"), " ",""), "'","")
				 If not Conn.Execute("Select ModelEname From KS_Channel Where ModelEname='" & ModelEname & "'").eof Then
				   Call KS.AlertHistory("系统已存在该目录名称，请重输!",-1)
				   Exit Sub
				 End If
				 If not Conn.Execute("Select ChannelTable From KS_Channel Where ChannelTable='" & ChannelTable & "'").eof Then
				   Call KS.AlertHistory("系统已存在该数据表，请重输!",-1)
				   Exit Sub
				 End If
				 Dim sChannelID:sChannelID=Conn.Execute("select Max(ChannelID) From KS_Channel")(0)+1
				 If sChannelID<100 Then sChannelID=sChannelID+100
				RS("ChannelID")    = sChannelID
				RS("BasicType")    = KS.ChkClng(KS.G("BasicType"))
				RS("ChannelTable") = ChannelTable
				RS("ModelEname")   =ModelEname
				RS("OrderID")   =MaxOrderID
			Else
			    OpName="修改"
			End If
				RS("ChannelName")= KS.G("ChannelName")
				RS("ItemName")   = ItemName
				RS("ItemUnit")   = KS.G("ItemUnit")
				RS("FsoFolder")  = KS.G("FsoFolder")
				RS("Descript")   = KS.G("Descript")
				if KS.ChkClng(RS("BasicType"))=1 Or KS.ChkClng(RS("BasicType"))=2 then  RS("CollectTF")=1
				RS("ChannelStatus") = KS.G("ChannelStatus")
				RS("WapSwitch")     = KS.ChkClng(KS.G("WapSwitch"))
				RS("FsoHtmlTF")     = KS.ChkClng(KS.G("FsoHtmlTF"))
				RS("StaticTF")      = KS.ChkClng(KS.G("StaticTF"))
				RS("FsoListNum")    = KS.ChkClng(KS.G("FsoListNum"))
				RS("UpfilesDir")    = KS.G("UpfilesDir")
				RS("UserUpfilesDir") = KS.G("UserUpfilesDir")
				RS("UpFilesTF")     = KS.G("UpFilesTF")
				RS("UserSelectFilesTF")=KS.G("UserSelectFilesTF")
				'If KS.G("UpfilesDir") <> "" Then Call KS.CreateListFolder(KS.Setting(3) & KS.G("UpfilesDir"))
				
				RS("UserUpFilesTF") = KS.G("UserUpFilesTF")
				'If KS.G("UserUpfilesDir") <> "" Then Call KS.CreateListFolder(KS.Setting(3) & KS.G("UserUpfilesDir"))
				
				RS("ThumbnailsConfig")=ThumbnailsConfig
	            RS("UserTF") = KS.ChkClng(KS.G("UserTF"))
				RS("UserEditTF")  = KS.ChkClng(KS.G("UserEditTF"))
				RS("UserClassStyle") = KS.ChkClng(KS.G("UserClassStyle"))
				RS("UpfilesSize") = KS.ChkClng(KS.G("UpfilesSize"))
				RS("AllowUpPhotoType") = KS.G("AllowUpPhotoType")
				RS("AllowUpFlashType") = KS.G("AllowUpFlashType")
				RS("AllowUpMediaType") = KS.G("AllowUpMediaType")
				RS("AllowUpRealType") = KS.G("AllowUpRealType")
				RS("AllowUpOtherType") = KS.G("AllowUpOtherType")
				RS("VerificCommentTF") = KS.G("VerificCommentTF")
				RS("LatestNewDay")     = KS.ChkClng(KS.G("LatestNewDay"))
				RS("CommentVF")    = KS.ChkClng(KS.G("CommentVF"))
				RS("CommentLen")   = KS.ChkClng(KS.G("CommentLen"))
				RS("CommentTemplate") = KS.G("CommentTemplate")
				RS("WapSearchTemplate")= KS.G("WapSearchTemplate")
				RS("InfoVerificTF") = KS.ChkClng(KS.G("InfoVerificTF"))
				RS("MaxPerPage")   = KS.ChkClng(KS.G("MaxPerPage"))
				RS("RefreshFlag")  = KS.ChkClng(KS.G("RefreshFlag"))
				RS("FsoContentRule")=KS.G("FsoContentRule")
				RS("FsoClassListRule")=KS.ChkClng(KS.G("FsoClassListRule"))
				RS("FsoClassPreTag")=KS.G("FsoClassPreTag")
				RS("ModelIco")=KS.G("ModelIco")
				RS("ModelShortName")=KS.G("ModelShortName")

				'会员积分
				RS("UserAddMoney") = KS.ChkClng(KS.G("UserAddMoney"))
				RS("UserAddPoint") = KS.ChkCLng(KS.G("UserAddPoint"))
				RS("UserAddScore") = KS.ChkClng(KS.G("UserAddScore"))
				RS("PubTimeLimit") = KS.ChkClng(KS.G("PubTimeLimit"))
				RS("ChargeType") = KS.ChkClng(KS.G("ChargeType"))
				RS("AnnexPoint") = KS.ChkClng(KS.G("AnnexPoint"))
				RS("DiggByVisitor")= KS.ChkClng(KS.G("DiggByVisitor"))
				RS("DiggByIP")     = KS.ChkClng(KS.G("DiggByIP"))
				RS("DiggRepeat")= KS.ChkClng(KS.G("DiggRepeat"))
				RS("DiggPerTimes")= KS.ChkClng(KS.G("DiggPerTimes"))
				RS.Update
				ChannelID=RS("ChannelID")
				ChannelTable=RS("ChannelTable")
				Dim BasicType:BasicType=RS("BasicType")
				RS.Close
				If BasicType=1 or BasicType=2 Or BasicType=3 Then
				    Dim Doc,Node,CDATASection
					set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
					Doc.async = false
					Doc.setProperty "ServerHTTPRequest", true 
					Doc.load(Server.MapPath(KS.Setting(3)&"Config/modelinputform.xml"))
					Set Node=Doc.documentElement.selectSingleNode("/inputform/model[@name='" & ChannelID & "']")
					 if not node is nothing then  Doc.DocumentElement.RemoveChild(Node)
					 Set Node=Doc.documentElement.appendChild(Doc.createNode(1,"model",""))
					 Node.attributes.setNamedItem(Doc.createNode(2,"name","")).text=channelid
					 Set   CDATASection   = Doc.createCDATASection(Request.Form("Content")) 
					 Node.appendChild   CDATASection 
					Doc.Save(Server.MapPath(KS.Setting(3)&"Config/modelinputform.xml"))
					Application(KS.SiteSN&"_Configmodelinputform")=empty
               End If
				
				
				
				If OpName="添加" Then
				 
				'建立新表
				dim sql
			    Select Case KS.ChkClng(KS.G("BasicType"))
			    Case 1
				sql="CREATE TABLE ["&ChannelTable&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&ChannelTable&" PRIMARY KEY,"&_
						"TID nvarchar(22),"&_
						"OTID nvarchar(22),"&_
						"OId int default 0,"&_
						"OrderId int default 0," &_
						"AvgScore float default 0,"&_
						"KeyWords nvarchar(255),"&_
						"TitleType nvarchar(30),"&_
						"Title nvarchar(255),"&_
						"FullTitle nvarchar(255),"&_
						"Intro ntext,"&_
						"ShowComment tinyint Default 0,"&_
						"TitleFontColor nvarchar(30),"&_
						"TitleFontType nvarchar(30),"&_
						"ArticleContent ntext,"&_
						"PageTitle ntext,"&_
						"Author nvarchar(30),"&_
						"Origin nvarchar(40),"&_
						"Rank nvarchar(10),"&_
						"Hits int Default 0,"&_
						"HitsByDay int Default 0,"&_
						"HitsByWeek int Default 0,"&_
						"HitsByMonth int Default 0,"&_
						"LastHitsTime datetime,"&_
						"AddDate datetime,"&_
						"ModifyDate datetime,"&_
						"JSID nvarchar(200),"&_
						"TemplateID nvarchar(255),"&_
						"WapTemplateID nvarchar(255)," &_
						"Fname nvarchar(200),"&_
						"RefreshTF tinyint default 0,"&_
						"Inputer nvarchar(50),"&_
						"PhotoUrl nvarchar(150),"&_
						"PicNews tinyint default 0,"&_
						"Changes tinyint default 0,"&_
						"Recommend tinyint Default 0,"&_
						"Rolls tinyint Default 0,"&_
						"Strip tinyint Default 0,"&_
						"Popular tinyint Default 0,"&_
						"Verific tinyint Default 0,"&_
						"Slide tinyint Default 0,"&_
						"Comment tinyint Default 0,"&_
						"IsTop tinyint Default 0,"&_
						"IsVideo tinyint Default 0,"&_
						"DelTF tinyint Default 0,"&_
						"PostID int Default 0,"&_
						"PostTable varchar(100),"&_
						"CmtNum int Default 0,"&_
						"IsSign tinyint Default 0,"&_
						"SignUser nvarchar(255),"&_
						"SignDateLimit tinyint Default 0,"&_
						"SignDateEnd datetime,"&_
						"Province nvarchar(100),"&_
						"City nvarchar(100),"&_
						"MapMarker nvarchar(255),"&_
						"InfoPurview tinyint Default 0,"&_
						"ArrGroupID nvarchar(100),"&_
						"ReadPoint int Default 0,"&_
						"ChargeType tinyint Default 0,"&_
						"PitchTime int Default 24,"&_
						"ReadTimes int Default 10,"&_
						"DividePercent int Default 0,"&_
						"SEOTitle varchar(255),"&_
						"SEOKeyWord ntext,"&_
						"SEODescript ntext,"&_
						"RelatedID int Default 0"&_
						")"
				Conn.Execute(sql)
				KS.ConnItem.Execute(sql)
				'添加索引
				Call AddIndex(ChannelTable, "[TID]", "[TID]")
				Call AddIndex(ChannelTable, "[Verific]", "[verific]")
				Call AddIndex(ChannelTable, "[deltf]", "[deltf]")
				Call AddIndex(ChannelTable, "[adddate]", "[adddate]")
				Call AddIndex(ChannelTable, "[hits]", "[hits]")
				'Call AddIndex(ChannelTable, "[specialid]", "[specialid]")
			 Case 2
				sql="CREATE TABLE ["&ChannelTable&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&ChannelTable&" PRIMARY KEY,"&_
						"Tid nvarchar(22),"&_
						"OTID nvarchar(22),"&_
						"OId int default 0,"&_
						"OrderId int default 0," &_
						"AvgScore float default 0,"&_
						"PicNum int default 0,"&_
						"Province varchar(100),"&_
						"City varchar(100),"&_
						"KeyWords nvarchar(255),"&_
						"Title nvarchar(255),"&_
						"showstyle tinyint default 0,"&_
						"pagenum int default 10,"&_
						"PhotoUrl nvarchar(255),"&_
						"PicUrls ntext,"&_
						"PictureContent ntext,"&_
						"Author nvarchar(30),"&_
						"Origin nvarchar(40),"&_
						"Rank nvarchar(10),"&_
						"LastHitsTime smalldatetime," &_
						"Hits int Default 0,"&_
						"HitsByDay int Default 0,"&_
						"HitsByWeek int Default 0,"&_
						"HitsByMonth int Default 0,"&_
						"AddDate smalldatetime,"&_
						"ModifyDate datetime,"&_
						"JSID nvarchar(200),"&_
						"TemplateID nvarchar(255),"&_
						"WapTemplateID nvarchar(255)," &_
						"Fname nvarchar(200),"&_
						"RefreshTF tinyint default 0,"&_
						"Inputer nvarchar(50),"&_
						"Recommend tinyint Default 0,"&_
						"Rolls tinyint Default 0,"&_
						"Strip tinyint Default 0,"&_
						"Popular tinyint Default 0,"&_
						"Verific tinyint Default 0,"&_
						"Slide tinyint Default 0,"&_
						"Comment tinyint Default 0,"&_
						"IsTop tinyint Default 0,"&_
						"Score int Default 0,"&_
						"MapMarker nvarchar(255),"&_
						"DelTF tinyint Default 0,"&_
						"PostID int Default 0,"&_
						"PostTable varchar(100),"&_
						"CmtNum int Default 0,"&_
						"InfoPurview tinyint Default 0,"&_
						"ArrGroupID nvarchar(100),"&_
						"ReadPoint int Default 0,"&_
						"ChargeType tinyint Default 0,"&_
						"PitchTime int Default 24,"&_
						"ReadTimes int Default 10,"&_
						"DividePercent int Default 0,"&_
						"SEOTitle varchar(255),"&_
						"SEOKeyWord ntext,"&_
						"SEODescript ntext"&_
						")"
				Conn.Execute(sql)
				KS.ConnItem.Execute(sql)
				'添加索引
				Call AddIndex(ChannelTable, "[TID]", "[TID]")
				Call AddIndex(ChannelTable, "[Verific]", "[verific]")
				Call AddIndex(ChannelTable, "[deltf]", "[deltf]")
				Call AddIndex(ChannelTable, "[adddate]", "[adddate]")
				Call AddIndex(ChannelTable, "[hits]", "[hits]")
				'Call AddIndex(ChannelTable, "[specialid]", "[specialid]")
				
			 Case 3
				sql="CREATE TABLE ["&ChannelTable&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&ChannelTable&" PRIMARY KEY,"&_
						"Tid nvarchar(22),"&_
						"OTID nvarchar(22),"&_
						"OId int default 0,"&_
						"OrderId int default 0," &_
						"AvgScore float default 0,"&_
						"KeyWords nvarchar(255),"&_
						"Title nvarchar(255),"&_
						"DownVersion nvarchar(50),"&_
						"DownLB nvarchar(100),"&_
						"DownYY nvarchar(100),"&_
						"DownSQ nvarchar(100),"&_
						"DownPT nvarchar(100),"&_
						"DownSize nvarchar(100),"&_
						"YSDZ nvarchar(100),"&_
						"ZCDZ nvarchar(100),"&_
						"JYMM nvarchar(100),"&_
						"PhotoUrl nvarchar(200),"&_
						"BigPhoto nvarchar(200),"&_
						"DownUrls ntext,"&_
						"DownContent ntext,"&_
						"Author nvarchar(50)," & _
						"Origin nvarchar(40),"&_
						"Rank nvarchar(10),"&_
						"LastHitsTime smalldatetime," &_
						"Hits int Default 0,"&_
						"HitsByDay int Default 0,"&_
						"HitsByWeek int Default 0,"&_
						"HitsByMonth int Default 0,"&_
						"AddDate datetime,"&_
						"ModifyDate datetime,"&_
						"JSID nvarchar(200),"&_
						"TemplateID nvarchar(255),"&_
						"WapTemplateID nvarchar(255)," &_
						"Fname nvarchar(200),"&_
						"RefreshTF tinyint default 0,"&_
						"Inputer nvarchar(50),"&_
						"Recommend tinyint Default 0,"&_
						"Rolls tinyint Default 0,"&_
						"Strip tinyint Default 0,"&_
						"Popular tinyint Default 0,"&_
						"Verific tinyint Default 0,"&_
						"Slide tinyint Default 0,"&_
						"Comment tinyint Default 0,"&_
						"IsTop tinyint Default 0,"&_
						"DelTF tinyint Default 0,"&_
						"PostID int Default 0,"&_
						"PostTable varchar(100),"&_
						"CmtNum int Default 0,"&_
						"InfoPurview tinyint Default 0,"&_
						"ArrGroupID nvarchar(100),"&_
						"ReadPoint int Default 0,"&_
						"ChargeType tinyint Default 0,"&_
						"PitchTime int Default 24,"&_
						"ReadTimes int Default 10,"&_
						"DividePercent int Default 0,"&_
						"SEOTitle varchar(255),"&_
						"SEOKeyWord ntext,"&_
						"SEODescript ntext"&_
						")"
				Conn.Execute(sql)
				'添加索引
				Call AddIndex(ChannelTable, "[TID]", "[TID]")
				Call AddIndex(ChannelTable, "[Verific]", "[verific]")
				Call AddIndex(ChannelTable, "[deltf]", "[deltf]")
				Call AddIndex(ChannelTable, "[adddate]", "[adddate]")
				Call AddIndex(ChannelTable, "[hits]", "[hits]")
			 Case 4
				sql="CREATE TABLE ["&ChannelTable&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&ChannelTable&" PRIMARY KEY,"&_
						"Tid nvarchar(22),"&_
						"OTID nvarchar(22),"&_
						"OId int default 0,"&_
						"OrderId int default 0," &_
						"AvgScore float default 0,"&_
						"KeyWords nvarchar(255),"&_
						"Title nvarchar(255),"&_
						"PhotoUrl nvarchar(255),"&_
						"FlashUrl varchar(255),"&_
						"FlashContent ntext,"&_
						"Author nvarchar(30),"&_
						"Origin nvarchar(40),"&_
						"Rank nvarchar(10),"&_
						"LastHitsTime smalldatetime," &_
						"Hits int Default 0,"&_
						"HitsByDay int Default 0,"&_
						"HitsByWeek int Default 0,"&_
						"HitsByMonth int Default 0,"&_
						"AddDate datetime,"&_
						"ModifyDate datetime,"&_
						"JSID nvarchar(200),"&_
						"TemplateID nvarchar(255),"&_
						"WapTemplateID nvarchar(255)," &_
						"Fname nvarchar(200),"&_
						"RefreshTF tinyint default 0,"&_
						"Inputer nvarchar(50),"&_
						"Recommend tinyint Default 0,"&_
						"Rolls tinyint Default 0,"&_
						"Strip tinyint Default 0,"&_
						"Popular tinyint Default 0,"&_
						"Verific tinyint Default 0,"&_
						"Slide tinyint Default 0,"&_
						"Comment tinyint Default 0,"&_
						"IsTop tinyint Default 0,"&_
						"Score int Default 0,"&_
						"MapMarker nvarchar(255),"&_
						"DelTF tinyint Default 0,"&_
						"PostID int Default 0,"&_
						"PostTable varchar(100),"&_
						"CmtNum int Default 0,"&_
						"InfoPurview tinyint Default 0,"&_
						"ArrGroupID nvarchar(100),"&_
						"ReadPoint int Default 0,"&_
						"ChargeType tinyint Default 0,"&_
						"PitchTime int Default 24,"&_
						"ReadTimes int Default 10,"&_
						"DividePercent int Default 0,"&_
						"SEOTitle varchar(255),"&_
						"SEOKeyWord ntext,"&_
						"SEODescript ntext"&_
						")"
				Conn.Execute(sql)
				KS.ConnItem.Execute(sql)
				'添加索引
				Call AddIndex(ChannelTable, "[TID]", "[TID]")
				Call AddIndex(ChannelTable, "[Verific]", "[verific]")
				Call AddIndex(ChannelTable, "[deltf]", "[deltf]")
				Call AddIndex(ChannelTable, "[adddate]", "[adddate]")
				Call AddIndex(ChannelTable, "[hits]", "[hits]")
				'Call AddIndex(ChannelTable, "[specialid]", "[specialid]")
				
			 End Select
				
				
				
				If KS.ChkClng(KS.G("BasicType"))=3 Then
				 Call KS.CreateListFolder(KS.Setting(3) & KS.G("UpfilesDir")&"DownPhoto/")
				 Call KS.CreateListFolder(KS.Setting(3) & KS.G("UpfilesDir")&"DownUrl/")
				End IF
				  
                 
				'  If Err<>0 Then
				'	Conn.RollBackTrans
				'	Call KS.AlertHistory("出错！出错描述：" & replace(err.description,"'","\'"),-1):response.end
				'  Else
				'	Conn.CommitTrans
				  'End If
				End If
				
				  Call KS.DelCahe(KS.SiteSN & "_selectallowclass")
				  Call KS.DelCahe(KS.SiteSN & "_selectclass")
				  Call KS.DelCahe(KS.SiteSN & "_classpath")
				  Call KS.DelCahe(KS.SiteSN & "_classnamepath")				     

				Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
				
				Call KSCLS.CreateModelField(ChannelID)     '初始化模型字段
				
                Application(KS.SiteSN & "_FieldGroupXml")=""
				Session("FromFile")="System/KS.Model.asp"
				KS.Die "<script>top.$.dialog.alert('KesionCMS系统提醒您：<br/>1、模型配置信息" & OpName & "成功；<br/>2、为了使配置生效，请及时更新缓存；',function(){ top.location.reload();})</script>"
			
		End Sub
	
		Sub DelColumn(TableName,ColumnName)
		On Error Resume Next
		Conn.Execute("Alter Table "&TableName&" Drop "&ColumnName&"")
		End Sub
		
		Sub DelTable(TableName,C)
			On Error Resume Next
			C.Execute("Drop Table "&TableName&"")
		End Sub
		
		Sub AddIndex(ByVal TableName, ByVal IndexName, ByVal ValueText)
			On Error Resume Next
			Conn.Execute("CREATE INDEX " & IndexName & " ON " & TableName & "(" & ValueText & ")")
		End Sub
		
		
		Sub ChannelDel()
		   On Error Resume Next
		  Dim ChannelID:ChannelID=KS.ChkClng(KS.G("ChannelID"))
		  Call DelTable(KS.C_S(ChannelID,2),Conn)
		  
		  '===============删除采集数据库里的相关字段和表=========================
		  If KS.C_S(ChannelID,6)="1" Or KS.C_S(ChannelID,6)="2" or KS.C_S(ChannelID,6)="5" Then  Call DelTable(KS.C_S(ChannelID,2),KS.ConnItem)
		  KS.ConnItem.Execute("Delete From KS_FieldItem Where ChannelID=" & ChannelID)
		  KS.ConnItem.Execute("Delete From KS_FieldRules Where ChannelID=" & ChannelID)
		  '=====================================================================
		 
		   '===============删除评论数据=====================================================
		  Dim TableXML:set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			TableXML.async = false
			TableXML.setProperty "ServerHTTPRequest", true 
			TableXML.load(Server.MapPath(KS.Setting(3)&"Config/commenttable.xml"))
		  Dim Node,id,isdefault:isdefault=KS.ChkClng(Request.Form("isdefault"))
		
		  For Each Node In TableXML.DocumentElement.SelectNodes("item")
		      Conn.Execute("Delete From " & node.selectsinglenode("tablename").text & " Where ChannelID=" & ChannelID)
		  Next
		  '=======================================================================
		  
		  Conn.Execute("Delete From KS_DownParam Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_DownSer Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Origin Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Channel Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Class Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Field Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_FieldGroup Where ChannelID=" & ChannelID)
		  
		  Call KS.DeleteFile(KS.Setting(3) &"config/fielditem/field_"&ChannelID&".xml")
		  Call KS.DeleteFile(KS.Setting(3) &"config/filtersearch/s"&ChannelID&".xml")
		  
		  '删除录入表单的模板
			Dim Doc,CDATASection
			set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/modelinputform.xml"))
			Set Node=Doc.documentElement.selectSingleNode("/inputform/model[@name='" & ChannelID & "']")
			if not node is nothing then  Doc.DocumentElement.RemoveChild(Node)
			Doc.Save(Server.MapPath(KS.Setting(3)&"Config/modelinputform.xml"))
			Application(KS.SiteSN&"_Configmodelinputform")=empty

		  
		  		 Call KS.DelCahe(KS.SiteSN & "_selectallowclass")
				 Call KS.DelCahe(KS.SiteSN & "_selectclass")
				 Call KS.DelCahe(KS.SiteSN & "_classpath")
				 Call KS.DelCahe(KS.SiteSN & "_classnamepath")
				 Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
			Session("FromFile")="System/KS.Model.asp"	 
		  Response.Write "<script>alert('模型删除成功!');top.location.href='../index.asp';</script>" 
		End Sub

		
		
			
End Class
%> 


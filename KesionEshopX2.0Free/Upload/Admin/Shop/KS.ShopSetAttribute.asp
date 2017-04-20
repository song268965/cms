<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New PaymentType
KSCls.Kesion()
Set KSCls = Nothing

Class PaymentType
        Private KS,ID,ChannelID,SearchParam,page,KeyWord,SearchType,ComeFrom,MaxPerPage,totalPut, CurrentPage,RS,I
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		  MaxPerPage=12
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
       Sub Kesion()
	    
		dim Price_Member,Price
		dim Param
		Param="where 1=1 "
		Price_Member=KS.G("Price_Member"):Price=KS.G("Price")
		Select case request("x")
		   case "e"
				if not IsNumeric(Price_Member) then Price_Member=0
				if not IsNumeric(Price) then Price=0
				conn.execute("Update KS_Product set Price_Member='" & Price_Member & "',Price='" & Price &"',TotalNum='"& KS.ChkClng(KS.G("TotalNum")) &"' where id="&  KS.ChkClng(KS.G("Shopid")) &"")
			    Response.Write("<script>top.$.dialog.alert('修改成功！');</script>")
		   case "a"
		        dim idstr,s_id
				idstr=KS.S("s_id")
				idstr=Split(idstr&"",",") 
				if UBound(idstr)>=0 then
					for i=0 to UBound(idstr)
						s_id=Trim(idstr(i))
						Price_Member=KS.G("Price_Member_"&s_id):Price=KS.G("Price_"&s_id)
						if not IsNumeric(Price_Member) then Price_Member=0
						if not IsNumeric(Price) then Price=0						
						conn.execute("Update KS_Product set Price_Member='" & Price_Member & "',Price='" & Price &"',TotalNum='"& KS.ChkClng(KS.G("TotalNum_"&s_id)) &"' where id="& Trim(idstr(i)) &"")
					next
					Response.Write("<script>top.$.dialog.alert('批量保存成功！');</script>")
				end if
			case "s"
						
		End Select
		I=0
		ChannelID=KS.ChkClng(KS.G("ChannelID"))
		If ChannelID=0 Then ChannelID=5
		ID = KS.G("ID"):If ID = "" Then ID = "0"
		KeyWord    = KS.G("KeyWord")    :  If KeyWord<>"" Then SearchParam=SearchParam & "&KeyWord=" & KeyWord
		SearchType = KS.G("SearchType") :  If SearchType<>"" Then  SearchParam=SearchParam & "&SearchType=" & SearchType
		If KeyWord<>"" Then
				Select Case SearchType
				  Case 0:Param = Param & " And (Title like '%" & KeyWord & "%')"
				  Case 4:Param = Param & " And ID Like '%" & KeyWord & "%'"
				End Select
		end if
		if ID <> "0" then  Param = Param & " And Tid In (" & KS.GetFolderTid(ID) & ")" 
		If Not KS.ReturnPowerResult(5, "M520075") Then  Call KS.ReturnErr(1, ""):Exit Sub   
         With Response
		   .Write "<!DOCTYPE html><html>"
			.Write"<head>"
			.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write"<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script src=""../../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
			.Write "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write"</head>"
			.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		End With
		%>
        <ul id='menu_top' class='menu_top_fixed'><li id='p7' class="parent"><a href='KS.ShopSetAttribute.asp'><span class="child"><i class="icon set"></i>直接批量改价</span></a></li><li id='p8' class="parent"><a href='../System/KS.ItemInfo.asp?action=SetAttribute&channelid=5'><span class="child"><i class="icon set"></i>按折扣批量改价</span></a></li> </ul>		
        <%
	
		 With KS
			  .echo ("<form action='KS.ShopSetAttribute.asp?x=s' name='searchform' method='get'>")
			  .echo ("<table height='43' border='0' width='100%' align='center'>")
			  .echo ("<tr><td><img src='../images/ico/search.gif' align='absmiddle'> <strong>快速搜索：</strong>")
			  .echo ("&nbsp;类型 <select name='searchtype'>")
			  If SearchType="4" Then .echo ("<option value=4 selected>商品编号</option>") Else .echo ("<option value=4>商品编号</option>")
			  If SearchType="0" Then .echo ("<option value=0 selected>商品名称</option>") Else .echo ("<option value=0>商品名称</option>")
			  .echo ("</select> <input type='text' class='textbox' title='关键字可留空' value='" & KeyWord &"' size='25' name='keyword'>")
			  .echo ("&nbsp;<input type='submit' class='button' value='开始搜索'><input type='hidden' value='" & ChannelID & "' name='channelid'></td>")
			  .echo ("</tr>")
			  .echo ("</table>")
			  .echo ("</form>")
			  End With 
		If KS.G("ComeFrom")="RecycleBin" Then
		     ShowChannelList 
		 Else
	         ShowClassList ChannelID,ID
		 End If
		 Page=KS.G("Page")
		 If KS.ChkClng(KS.G("Page"))=0 Then
				 CurrentPage = 1
		 else
		 		 CurrentPage=KS.ChkClng(KS.G("Page"))
		 End If
%>		
		<script>
        function ShopEdit(id,Q_str){
			location.href='KS.ShopSetAttribute.asp?x=e&'+Q_str+'&Shopid='+id +'&Price_Member='+ $("input[name=Price_Member_"+ id +"]").val() +'&Price='+$("input[name=Price_"+ id +"]").val()+'&TotalNum='+$("input[name=TotalNum_"+ id +"]").val()
		}
		function ClassToggle(f)
		{
		  setCookie("classExtStatus",f)
		  $('#classNav').toggle('slow');
		  $('#classOpen').toggle('show');
		}

        </script> 
        <div class="pageCont2 mt20">
		<form name="form1" method="post" action="?x=a<%=SearchParam%>">
        <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr align="center"  class="sort"> 
			<td width="87"><strong>编号</strong></td>
			<td><strong>商品名称</strong></td>
			<td><strong>商城价</strong></td>
			<td><strong>参考价</strong></td>
			<td><strong>库存数量</strong></td>
			<td><strong>管理操作</strong></td>
		  </tr>
		  <%dim orderid
		  Set RS = Server.CreateObject("ADODB.RecordSet")
		  RS.Open "select * from KS_Product "& Param &"  order by ID desc", conn, 1, 1
		  If RS.Eof And RS.Bof Then
	 			Response.Write "<tr><td colspan=""6"" height=""25"" align=""center"" class=""tdbg"">还没有记录！</td></tr>"
		  Else
                     
					  totalPut = RS.RecordCount
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							End If
							Call showContent
  		  End If
		  
%>
 <tr>
 <td colspan=6>
 <input name="Page" type="hidden" id='Page' value='<%=Page%>'>
 <input name="ChannelID" type="hidden" id='ChannelID' value='<%=ChannelID%>'>
 <input name="ID" type="hidden" id='ID' value='<%=ID%>'>
 <br /><input type='submit' name="Submit" value='批量保存' class='button'> </td>
 </form>
 </tr>
 <tr>
 <td colspan=7 ><% Call KS.ShowPage(totalput, MaxPerPage, "",CurrentPage,true,true) %></td>
 
 </tr>
</table>
</div>

		</body>
		</html>
<%End Sub


Sub ShowContent()
	Do While Not RS.EOF
	%>
				<tr  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
				  <td  class="splittd" height="25" align="center"><%=rs("ID")%> <input name="s_id" type="hidden" id="s_id" value="<%=rs("ID")%>"></td>
				  <td  class="splittd" align="left"><%=rs("title")%></td>
				  			  
				  <td  class="splittd" align="center"><input style="text-align:center"  name="Price_Member_<%=rs("ID")%>" type="text" class="textbox" id="Price_Member" value="<%=KS.GetPrice(rs("Price_Member"))%>"  size="8">
				  </td>
				  <td class="splittd" align="center"><input style="text-align:center" name="Price_<%=rs("ID")%>" type="text" class="textbox" id="Price" value="<%=KS.GetPrice(rs("Price"))%>"  size="10">
				  </td>
				  <td class="splittd" align="center">
				 <input style="text-align:center" name="TotalNum_<%=rs("ID")%>" type="text" class="textbox" id="TotalNum" value="<%=rs("TotalNum")%>" size="10" onKeyUp="value=value.replace(/\D/g,'')" onafterpaste="value=value.replace(/\D/g,'')">
				  </td>	
				  <td class="splittd" align="center"><input  class="button"  type="button"  onclick="ShopEdit(<%=rs("ID")%>,'<%=ks.QueryParam("x,Price,Price_Member,shopid,TotalNum")%>');" value=" 修改 "></td>
				</tr>
			  
    		<% I = I + 1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		Loop
		RS.Close
End Sub

   Sub ShowClassList(ChannelID,ID)
		 If KS.S("ComeFrom")<>"" Then Exit Sub
		 
		 With KS
		 '============增加记忆功能=======================================
		 Dim ExtStatus,CloseDisplayStr,ShowDisplayStr,classExtStatus
		 classExtStatus=request.cookies("classExtStatus")
		 if classExtStatus="" Then classExtStatus=1
		 If classExtStatus=1 Then 
		  ExtStatus=2 :CloseDisplayStr="display:none;":ShowDisplayStr=""
		 Else 
		  ExtStatus=1 :CloseDisplayStr="":ShowDisplayStr="display:none;"
		 End If
		 '=========================================================----
		 .echo "<div id='classOpen' onclick=""ClassToggle("& ExtStatus& ")"" style='" & CloseDisplayStr &"cursor:pointer;text-align:center;position:absolute; z-index:2; left: 0px; top: 38px;' ><img src='../images/kszk.gif' align='absmiddle'></div>"
		 .echo "<div id='classNav' style='" & ShowDisplayStr &"position:relative;height:auto;_height:30px;line-height:30px;margin:5px 1px;'>"
		 .echo "<ul><div style='cursor:pointer;text-align:center;position:absolute; z-index:1; right: 0px;'  onclick=""ClassToggle(" & ExtStatus &")""> <img src='../images/mk_del.png' align='absmiddle'></div>"
		
		Dim P,RSC,Img,j,N,I,XML,Node
		P=" where ClassType=1 and ChannelID=" & ChannelID
		If ID=0 Then
		  P=P & " And tj=1"
		 Img="domain.gif"
		Else
		 P=P & " And TN='" & ID & "'"
		 Img="Smallfolder.gif"
		End If

		 on error resume next
		 Dim ParentID:ParentID = conn.Execute("Select TN From KS_Class  Where ID='" & ID & "'")(0)

		Set RSC=Conn.Execute("select id,foldername,adminpurview from ks_class " & P& " order by root,folderorder")
		If Not RSC.Eof Then 
		 Set XML=.RsToXml(RSC,"row","xmlroot")
		 RSC.Close:Set RSC=Nothing
		 If IsObject(XML) Then
		   If ID<>"0" Then
		    .echo "<a href='?ChannelID=" & ChannelID & "&ID=" & ParentID & "'><i class='icon back'></i></a> <div class='clear'></div>"
		   End if
		   For Each Node In XML.DocumentElement.SelectNodes("row")
		    If KS.C("SuperTF")=1 or KS.FoundInArr(Node.SelectSingleNode("@adminpurview").text,KS.C("GroupID"),",") or Instr(KS.C("ModelPower"),KS.C_S(ChannelID,10)&"1")>0 Then 
		    .echo "<li><i class=""icon folder""></i><a href='?ChannelID=" & ChannelID & "&ID=" & Node.SelectSingleNode("@id").text & "' title='" & Node.SelectSingleNode("@foldername").text & "'>" & .Gottopic(Node.SelectSingleNode("@foldername").text,8) & "(<span style='color:#ff6600'>" &conn.execute("select count(1) from " & KS.C_S(ChannelID,2)& " where deltf=0 and tid in(select id from ks_class where ts like '%" & Node.SelectSingleNode("@id").text &",%')")(0) &"</span>)</a></li>"
		    End If
		   Next
		 End If
		Else
		  If err Then
		   .echo "<i class='icon folder'></i>请先<a href='#' onclick=""location.href='KS.Class.asp?Action=Add&ChannelID=" & ChannelID & "';$(parent.document).find('#BottomFrame')[0].src='Post.Asp?ButtonSymbol=Go&OpStr=" & Server.URLEncode("栏目管理 >> <font color=red>添加栏目</font>") & "';"">添加栏目</a>"
		  Else
		   .echo "<a href='?ChannelID=" & ChannelID & "&ID=" & ParentID & "'><i class='icon back'></i></a> "
		   End If
		End If
		 .echo "</ul></div>"
		 .echo "<div style=""clear:both""></div>"
		 End With
		End Sub
		
		Sub ShowChannelList()
		  With KS
			 '============带记忆功能=======================================
			 Dim ExtStatus,CloseDisplayStr,ShowDisplayStr,classExtStatus
			 classExtStatus=request.cookies("classExtStatus")
			 if classExtStatus="" Then classExtStatus=1
			 If classExtStatus=1 Then 
			  ExtStatus=2 :CloseDisplayStr="display:none;":ShowDisplayStr=""
			 Else 
			  ExtStatus=1 :CloseDisplayStr="":ShowDisplayStr="display:none;"
			 End If
			 '=========================================================----
			 .echo "<div id='classOpen' onclick=""ClassToggle("& ExtStatus& ")"" style='" & CloseDisplayStr &"cursor:pointer;text-align:center;position:absolute; z-index:2; left: 0px; top: 2px;' ><img src='../images/kszk.gif' align='absmiddle'></div>"
			 .echo "<div id='classNav' style='" & ShowDisplayStr &"position:relative;height:auto;_height:30px;line-height:30px;margin:5px 1px;border:1px solid #DEEFFA;background:#F7FBFE'>"
			 .echo "<ul><div style='cursor:pointer;text-align:center;position:absolute; z-index:1; right: 0px;'  onclick=""ClassToggle(" & ExtStatus &")""> <img src='../images/close.gif' align='absmiddle'></div>"
			 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Dim ModelXML,Node
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel")
				 if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"10" Then
				   .echo "<li style='margin:5px;float:left;width:100px'><i class='icon mainer'></i><a href='?ChannelID=" & Node.SelectSingleNode("@ks0").text & "&ComeFrom=RecycleBin' title='" & Node.SelectSingleNode("@ks1").text & "'>" & .Gottopic(Node.SelectSingleNode("@ks1").text,8) & "(<span style='color:red'>" & Conn.Execute("Select Count(ID) From " & Node.SelectSingleNode("@ks2").text & " Where Deltf=1")(0) & "</span>)</a></li>"
			    End If
			next
			.echo "</ul></div>"
			.echo "<div style=""clear:both""></div>"
         End With
		End Sub
		
End Class
%> 

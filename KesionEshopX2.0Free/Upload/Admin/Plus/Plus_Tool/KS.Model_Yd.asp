<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../../Include/Session.asp"-->
<%
Response.Buffer=true
Response.CharSet="utf-8"
Server.ScriptTimeout=9999999

'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Import
KSCls.Kesion()
Set KSCls = Nothing

Class Import
        Private KS,KSCls,ChannelID,IConnStr,Iconn,tempField
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		 Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Sub Kesion()
		 If KS.S("Action")="testsource" Then
		   Call testsource()
		   Exit Sub
		 End If
		 With KS
			.echo "<!DOCTYPE html><html>"
			.echo "<title>基本参数设置</title>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.echo "<link href=""../../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script src=""../../../KS_Inc/common.js"" language=""JavaScript""></script>"
			.echo "<script src=""../../../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			.echo "</head>"
			.echo "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
		  
		  select case request("action")
		   case "Step2" Step2
		   case "Step3" Step3
		   case else
		     call step1
		  end select
		  	.echo "</body>"
			.echo "</html>"
        End With
	   End Sub
	   
	   
	   '
	   Sub Step1()
	   	%>
		<script>
        	 $(function() {
				$("#channelid").change(function(){
				    location.href='?Action=Step1&bstr='+$(this).val()

				});
				$("#channelid_n").change(function(){
					if($(this).val()!=0){
						var cid=$(this).val()
						$("#tClassID").html($("#SClass_"+cid).html())
						$("#ttClassID").html($("#SClass_"+cid).html())
					}
				});	
			 });
			 function CheckydForm(){
			 	if ($("#channelid").val()==0){
					alert("请选择源模型!");
					return false;
				}
				if ($("#channelid_n").val()==0){
					alert("请选择目标模型");
					return false;
				}
			 }
        </script>
		<%
		dim bstr,BasicType,wherestr	
		 bstr=split(KS.G("bstr")&"","_")
		 if Ubound(bstr)>0 then
		 	ChannelID=KS.ChkClng(bstr(0))
			BasicType=KS.ChkClng(bstr(1))
		 else
		 	ChannelID=0	
		 end if
		    
		 With KS
			.echo "      <div class='tabTitle mt20'>"
			.echo "      模型数据移动"
			.echo "      </div>"
			.echo "      <div class='pageCont2'>"
			.echo "<form action=""?Action=Step2"" method=""post"" name=""DownParamForm"" onSubmit=""return(CheckydForm())"">"
			.echo "  <table width=""100%"" border=""0"" align=""center"" cellspacing=""1"" class=""ctable"">"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>要移动的模型</strong></td>"
			.echo "      <td>"
			.echo "<select id='channelid' name='channelid'>"
			.echo " <option value='0'>---请选择源模型---</option>"
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node,zs
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1 and @ks6<9]")
				'if =Node.SelectSingleNode("@ks0").text then
				zs=Conn.Execute("SELECT COUNT(ID) FROM "& Node.SelectSingleNode("@ks2").text &" ")(0) 
			   .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"_"& Node.SelectSingleNode("@ks6").text  &"'"
			   if ChannelID=KS.ChkClng(Node.SelectSingleNode("@ks0").text) then .echo "selected=""selected"""
			   .echo">" & Node.SelectSingleNode("@ks1").text & "(" & Node.SelectSingleNode("@ks2").text & ")[共"&zs& Node.SelectSingleNode("@ks4").text &"]</option>"
			   ' end if
			next
			.echo "</select> 移动到<font size=""+1"">→</font> "		
			.echo "<select id='channelid_n' name='channelid_n'>"
			.echo " <option value='0'>---请选择目标模型---</option>"
			if BasicType<>0 then
				If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				if BasicType<>0 then wherestr="@ks6="&BasicType else wherestr="(@ks6=3 or @ks6=1 or @ks6=5)"
				For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks0!=" & channelid &" and @ks21=1 and "&wherestr&"]")
				   .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "(" & Node.SelectSingleNode("@ks2").text & ")</option>"
				next
			end if
			.echo "</select>"
			.echo "     </td>"
			.echo "    </tr>"

			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>移动数据</strong></td>"
            .echo "      <td>"
			%>
            
            <input type="radio" name="dtype" value="1" checked onclick="$('#dclass').show();$('#did').hide();$('.mclass').show();$('#option').show();" />按指定栏目移动指定栏目  <input type="radio" name="dtype" onclick="$('#dclass').hide();$('#did').show();$('#option').show();" value="2" />按ID号移动 <input type="radio" name="dtype" onclick="$('#dclass').show();$('#did').hide();$('.mclass').hide();$('#option').hide();" value="3" />指定栏目移动到新建模型表
            <br /><br />
            <div id="dclass" style=" width:700px;">
                <div style="float:left;">
                <select name='BatchClassID' size='2' multiple style='height:250px;width:300px;'>
				<%
				if ChannelID<>0 then Response.Write KS.LoadClassOption(ChannelID,false)
				%>
                </select><br> <br> 
                <input type='button' name='Submit' value='选定所有' class='button' onclick='SelectAll()'>        <input type='button' class='button' name='Submit' value='取消所选' onclick='UnSelectAll()'> 
				
                </div>
                <div class='mclass' style="float:left; padding:20px;">
                    移动到&gt;&gt;
                </div>
                <div class='mclass' style="float:left;">
                    <select name='tClassID'  id="tClassID" size='2' style='height:250px;width:300px;'></select> 
                </div>
            </div>
            
            <div id="did" style="display:none;">
            	开始ID号<input type="text" class="textbox" name="S_ID" value="1" /> 结束ID号<input class="textbox" type="text" name="E_ID" value="1000" /><br />
                移动到栏目↓<br />
                 <select name='ttClassID' id="ttClassID" size='2' style='height:250px;width:300px;'></select> 
				 
            </div>
            <div  style="display:none">
            <%dim ChannelID_t
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1]")
			%>
				 <%ChannelID_t=KS.ChkClng(Node.SelectSingleNode("@ks0").text)%>
                 <select name='SClass_<%=ChannelID_t%>' id='SClass_<%=ChannelID_t%>' size='2' style='height:360px;width:300px;'>
				 	<%=KS.LoadClassOption(ChannelID_t,false)%>
                 </select>
			<%next%>
			</div>
			<div id="option">
			<br style="clear:both"/><br/><label><input type="checkbox" name="refname" value="1" checked="checked" />自动更新目标生成的HTML文件名,以防止文件名重复存在。 </label>当目标数据表为空时，不建议打勾。
			<br/><br/><label><input type="checkbox" name="retemplate" value="1" checked="checked" />自动绑定模板为目标栏目的默认文档模板。 </label>
			</div>
			<%
			.echo "</td>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo " <div style='text-align:center;padding:20px'><input type='submit' value=' 执行移动 ' class='button' name='button1'></div>"
			.echo "</form>"
			.echo "</div>"
			End With
			%>
			<div class="attention">
		<font color=red><strong>说明：</strong><br/>此功能是针对前期网站没有架构规划完善，原主模型的数据量极大，导致系统运行缓慢，利用自定义模型分表来分担数据而开发的插件。即将原主模型的数据移到部分到新的自定义模型从而达到减少主模型数据量的目的。数据量不大的情况，不建议使用。</font>
		</div>

			<%
		End Sub
		Sub Step2()
			Dim Cid,CBasic,i,FolderTidList
			Dim channelid_arr:channelid_arr=split(KS.G("channelid")&"","_")
			if Ubound(channelid_arr)>0 then
				Cid=KS.ChkClng(channelid_arr(0))
				CBasic=KS.ChkClng(channelid_arr(1))
			else
				 KS.AlertHintScript "请选择源模型!"
				 Response.End()
			end if
			Dim channelid_n:channelid_n=KS.ChkClng(KS.G("channelid_n"))
			If channelid_n=0 Then 
				 KS.AlertHintScript "请选择目标模型!"
				 Response.End()
			End If
			if cid=channelid_n then 
				 KS.AlertHintScript "请选择不同的两个模型!"
				 Response.End()
			end if
			
			'判断表结构是否相同
			dim field1:field1=GetTableField(KS.C_S(Cid,2))
			dim field2:field2=GetTableField(KS.C_S(channelid_n,2))
			field1=lcase(field1)
			field2=lcase(field2)
			if field1<>field2 then
				dim nofield,ii,fieldarr:fieldarr=split(field1,",")
				for ii=0 to ubound(fieldarr)
				  if KS.FoundInarr(field2,fieldarr(ii),",")=0 then
					if nofield="" then
					  nofield=fieldarr(ii)
					else
					  nofield=nofield &"," & fieldarr(ii)
					end if
				  end if 
				next
				if nofield<>"" then 
				 KS.AlertHintScript "目标模型不存在以下字段“" & nofield & "”，请先建立才能移动!"
				End If
		    End If	
			
            Dim dtype:dtype=KS.G("dtype")
			Dim BatchClassID:BatchClassID=Replace(KS.G("BatchClassID")," ","")
		 	Dim tClassID:tClassID=KS.G("tClassID")
			if tClassID="" Then TclassID=KS.G("ttClassID")
			If BatchClassID="" and dtype<>"2" Then 
		    	KS.AlertHintScript "请选择要移动栏目！"
				Response.End()
		    End if
			
			If tClassID="" and dtype<>"3" Then
			    KS.AlertHintScript "请先择目标栏目！"
				Response.End()
			End If
			
			BatchClassID=Split(BatchClassID,",")
			For i=0 To Ubound(BatchClassID)
				 If FolderTidList="" Then
				 	FolderTidList=GetFolderTid(BatchClassID(i))
				 Else
					FolderTidList=FolderTidList &","&GetFolderTid(BatchClassID(i))
				 End If
		    Next
			dim param
			if dtype="2" then
			  param=" where id>=" &KS.ChkClng(Request("s_id")) &" and id<=" & KS.ChkClng(Request("e_id"))
			else
			  param="  Where tid in(" & foldertidlist &")"
			end if
            If DataBaseType=1 Then  CONN.execute("SET IDENTITY_INSERT [" & KS.C_S(channelid_n,2) & "] ON")	
				
			Dim Field_Arr:Field_Arr=Split(GetTableField(KS.C_S(Cid,2)),",")
			Dim RS1:Set RS1=Server.CreateObject("adodb.recordset")
			Dim RS2:Set RS2=Server.CreateObject("adodb.recordset")
			RS1.Open "select * From " & KS.C_S(Cid,2) & param & " order by id desc",conn,1,1
			RS2.Open "select * from " & KS.C_S(channelid_n,2) & " Where 1=0",conn,1,3
			dim total:total=RS1.Recordcount
			if total=0 then
			 'KS.AlertHintScript "您选择的源模型没有数据！"
			' response.End()
			end if
			response.write "<div class=""attention"">" &vbcrlf
			response.flush
			response.write "<li><font color=blue>共需要移动" & total & "条记录！</font></li>"
			response.flush
			dim nn:nn=0
			dim tipsNum:tipsNum=10
			If total>100000 Then 
			 tipsNum=1000
			ElseIf Total>50000 Then
			 tipsNum=500
			ElseIf Total>10000 Then
			 tipsNum=200
			ElseIf Total>5000 Then
			 tipsNum=100
			End If
			
			Do While Not RS1.Eof
			   nn=nn+1
			   RS2.AddNew
			     For II=0 To Ubound(Field_Arr)
					   IF dtype="3" THEN  '直接按栏目移动到模型
						 RS2(trim(Field_Arr(ii)))=RS1(trim(Field_Arr(ii)))
					   ELSE
						   if Field_Arr(ii)="id" then
						   elseif Field_Arr(ii)="tid" then
							RS2("tid")=tClassID
						   Elseif request("retemplate")=1 and (Field_Arr(ii)="templateid" or Field_Arr(ii)="waptemplateid") then
							 if Field_Arr(ii)="templateid" then
								RS2(trim(Field_Arr(ii)))=KS.C_C(TclassID,5)
							 else
								RS2(trim(Field_Arr(ii)))=KS.C_C(TclassID,22)
							 end if
						   Else
							RS2(trim(Field_Arr(ii)))=RS1(trim(Field_Arr(ii)))
						   end if
					   END IF
				 Next
			   RS2.Update
			   If Request("refname")="1" and dtype<>"3" Then
			    RS2.MoveLast
				dim newfname:newfname=rs2("id")&".html"
				RS2("Fname")=newfname
				RS2.Update
				Conn.Execute("Update KS_ItemInfo Set channelid=" &channelid_n &",fname='"&newfname & "',Tid='" & TclassID & "' Where channelID=" & cid & " and infoid=" &  rs1("id"))
			   End If
			   if nn mod tipsNum =0 then
			   response.write "<li>已移动" & nn & "条数据!</li>"
			   response.flush
			   end if
			RS1.MoveNext
			Loop
			RS1.Close
			Set RS1=Nothing
			RS2.Close
			Set RS2=Nothing
			
			conn.Execute("Delete  From " & KS.C_S(Cid,2) & " " & param)
			
			IF dtype="3" THEN
			   'Conn.Execute("Update KS_ItemInfo Set channelid=" & channelid_n &" " & param) 
			   Conn.Execute("Update KS_Class Set ChannelID=" & channelid_n & " where id in(" & foldertidlist &")")
			Else
				if request("refname")<>"1" then
				 Conn.Execute("Update KS_ItemInfo Set channelid=" & channelid_n &",Tid='" & tClassID & "' " & param) 
				end if
			End If
			If DataBaseType=1 Then  CONN.execute("SET IDENTITY_INSERT [" & KS.C_S(channelid_n,2) & "] OFF")
			
			'更新标签所属模型ID
			 Dim ClassIDArr:ClassIDARR=split(replace(FolderTidList&"","'",""),",")
			 Dim n
			 DIM RSL:Set RSL=Server.CreateObject("Adodb.recordset")
			For n=0 To Ubound(ClassIDArr)
			  RSL.Open "select * from ks_label where labelcontent like '%classid=""" &ClassIDArr(n) &"""%'",conn,1,3
			  Do While Not RSL.Eof
			    Dim LabelContent:LabelContent=Replace(RSL("LabelContent"),"modelid=""" & Cid& """","modelid=""" & channelid_n& """")
				RSL("LabelContent")=LabelContent
				RSL.Update
			  RSL.MoveNext
			  Loop
			  RSL.Close
			Next
			 
			 Call KS.DelCahe(KS.SiteSN & "_selectallowclass")
			 Call KS.DelCahe(KS.SiteSN & "_selectclass")
			 Call KS.DelCahe(KS.SiteSN & "_classpath")
			 Call KS.DelCahe(KS.SiteSN & "_classnamepath")
				 	
			response.write "<li>已移动" & total & "条数据!</li>"
			 response.flush
			response.write "<li style='color:blue'>恭喜，所有文档移动成功！</li>"
			response.write "</div>"
			response.write "<div style='text-align:center'><input type='button' onclick='location.href=""KS.Model_Yd.asp"";' class='button' value=' 返 回 '/></div>"
		End Sub



		Function GetFolderTid(FolderID)
			Dim I,Tid,SQL
			Dim RS:Set RS=Conn.Execute("Select ID From KS_Class Where DelTF=0 AND TS LIKE '%" & FolderID & "%'")
			 If RS.EOF Then	 GetFolderTid="'0'":RS.Close:Set RS=Nothing:Exit Function
			 SQL=RS.GetRows(-1):RS.Close:Set RS = Nothing
             For I=0 To Ubound(SQL,2)
				  Tid = Tid & "'" & Trim(SQL(0,I)) & "',"
			 Next
			Tid = Left(Trim(Tid), Len(Trim(Tid)) - 1) '去掉最后一个逗号
			GetFolderTid = Tid
		End Function
		Sub testsource()
			response.cachecontrol="no-cache"
			response.addHeader "pragma","no-cache"
			response.expires=-1
			response.expiresAbsolute=now-1
			Response.CharSet="utf-8"
			on error resume next
		   dim str:str=unescape(request("str"))
		   If KS.G("DataType")="1" or KS.G("DataType")="2" Then str=LFCls.GetAbsolutePath(str)
		   dim tconn:Set tconn = Server.CreateObject("ADODB.Connection")
			tconn.open str
			If Err Then 
			  Err.Clear
			  Set tconn = Nothing
			  KS.Echo "false"
			else
			  KS.Echo "true"
			end if
		End Sub
		
		Sub OpenImporIConn()
				   if not isobject(IConn) then
					on error resume next
					Set IConn = Server.CreateObject("ADODB.Connection")
					IConn.open IConnStr
					If Err Then 
					  Err.Clear
					  Set IConn = Nothing
					  Response.Write "<script>alert('数据源连接失败,请检查数据库连接!');history.back();</script>"
					  response.end
					end if
				   end if		
		End Sub
       '**************************************************
		'过程名：ShowChird
		'作  用：显示指定数据表的字段列表
		'参  数：无
		'**************************************************
		Function GetTableField(dbname)
				dim tempField:tempField=""
					dim rs:Set rs=conn.OpenSchema(4)
					Do Until rs.EOF or rs("Table_name") = trim(dbname)
						rs.MoveNext
					Loop
			
					Do Until rs.EOF or rs("Table_name") <> trim(dbname)
					  if tempField="" then
					   tempField=rs("column_Name")
					  else
					  tempField=tempField & ","&rs("column_Name")
					  end if
					  rs.MoveNext
					loop
				    rs.close:set rs=nothing
			   GetTableField=lcase(tempField)
		End Function	
		
		
		
		
		


End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
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
		
		  if request("action")="geturls" then
		       response.cachecontrol="no-cache"
				response.addHeader "pragma","no-cache"
				response.expires=-1
				response.expiresAbsolute=now-1
				Response.CharSet="utf-8"
		       dim folderpath:folderpath=KS.G("folderpath")
			   dim subfolder:subfolder=ks.chkclng(ks.g("subfolder"))
			   call geturls(folderpath,subfolder)
			   ks.die("")
		  end if
		 
		 With KS
			.echo "<html>"
			.echo "<title>批量添加文章</title>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.echo "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script src=""../../KS_Inc/common.js"" language=""JavaScript""></script>"
			.echo "<script src=""../../KS_Inc/jquery.js"" language=""JavaScript""></script>"
           %>
		    <script type="text/javascript">
			function check(){
			  if ($("#channelid option:selected").val()==0){
			    alert('请选择模型!');
				return false;
			  }
			  if ($("#ClassID option:selected").val()==0){
			    alert('请选择栏目!');
				return false;
			  }
			  <%if request("action")="addTxt" then%>
			  if ($("#TxtDir").val()==''){
			    alert('请输入Txt文件所在的目录!');
				$("#TxtDir").focus();
				return false;
			  }
			  if ($("#AddressUrls").val()==''){
			    alert('没有找到Txt文件，请先点[获取Txt文件列表]!');
				return false;
			  }
			  <%end if%>
			  return true;
			}
			function getClass(channelid){
			$.get('../../plus/ajaxs.asp',{action:'GetClassOption',channelid:channelid},function(data){
			  $("#ClassID").empty();
			  $("#ClassID").append(unescape(data));
			 });
		   }
		    function getAdd(v){
			  var str='';
			  for(var i=1;i<=v;i++){
               str+='<tr><td height="35" align="right"><font color="#993300">第'+i+'篇</font>&nbsp;题目标题：</td>';
			   str+=' <td><input maxlength="80" type="text" class="textbox" style="width:300px;" id="title'+i+'" name="title'+i+'"><font color=red>*</font></td></tr>';
			   str+='<tr><td height="35" align="right">文章内容：</td>';
               str+=' <td><textarea name="content'+i+'" cols="60" id="content'+i+'" rows="5"></textarea></td></tr>';
			   str+="<tr><td colspan='2'><hr></td></tr>";
			  }
			  $("#showadd").empty().append(str);
			}
			$(function(){
			 getAdd(10);
			});
			</script>
		   <%
			.echo "</head>"
			.echo "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
		    .echo "      <div class='topdashed sort' style='text-align:left'>"
			.echo "      <a href='KS.ArticleAdd.asp'>批量添加(批量录入)</a> <a href='KS.ArticleAdd.asp?action=addTxt'>批量添加(TXT导入)</a>"
			.echo "      </div>"
			
		    if request("action")="dosave" then
			  dosave
			elseif request("action")="dotxtsave" then
			  dotxtsave
			elseif request("action")="addTxt" then
			  addTxt
			else
			  add
		    end if
		  	.echo "</body>"
			.echo "</html>"
        End With
	   End Sub
	   
	   sub addHead()
	     with KS
			.echo "  <table width=""100%"" border=""0"" align=""center"" cellspacing=""1"" class=""ctable"">"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>要添加的模型</strong></td>"
			.echo "      <td><select id='channelid' name='channelid' onchange=""getClass(this.value)"">"
			.echo " <option value='0'>---请选择目标模型---</option>"
	
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1 and @ks6=1]")
			   .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "(" & Node.SelectSingleNode("@ks2").text & ")</option>"
			next
			.echo "</select>"			
			.echo "     </td>"
			.echo "    </tr>"
			%>
			<tr class='tdbg'>
			<td width="150" height="30" class='clefttitle' align='right'><strong>批量添加到栏目</strong></td>
            <td><select name="ClassID" id="ClassID">
			 <option value=''>=请选择模型=</option>
			</select></td>
			 </tr>
	  <%
		 End With
	   end sub
	   
	   
	   Sub geturls(byval folderpath,subfolder)
		  if right(folderpath,1)<>"/" then folderpath=folderpath & "/"
		  dim FsoObj,FolderObj,FileObj,FsoItem,fext
		  Set FsoObj = KS.InitialObject(KS.Setting(99))
		  If Not FsoObj.FolderExists(Server.MapPath(folderpath)) Then
		   ks.die "errorfolder" 
		  end if
		  Set FolderObj = FsoObj.GetFolder(Server.MapPath(folderpath))
		  Set FileObj = FolderObj.Files
		  For Each FsoItem In FileObj
		      fext=Mid(FsoItem.name, InStrRev(FsoItem.name, ".")) '分离出扩展名
			  if (lcase(fext)=".txt") then
			  KS.Echo "<option>" & folderpath & FsoItem.name &"</option>"
			  end if
		  Next
		  '子目录
		  If subfolder=1 then
			  if FolderObj.subfolders.count>0 then
				For Each FsoItem In FolderObj.subfolders
				  call geturls (folderpath & fsoitem.name,subfolder)
				Next
			  end if
		  end if
		End Sub
	   
       
	   
	   sub addTxt()
	     response.write "<form action=""?Action=dotxtsave"" method=""post"" name=""DownParamForm"">"
	      addHead
		  %>
		  <tr class='tdbg'>
			<td height="30" class='clefttitle' align='right'><strong>录入员：</strong></td>
            <td><input type="text" name="inputer" class="TextBox" value="<%=KS.C("AdminName")%>"/></td>
		 </tr>
		  <tr class='tdbg'>
			<td height="30" class='clefttitle' align='right'><strong>TXT文件所在目录：</strong></td>
            <td><input type="text" name="TxtDir" id="TxtDir" class="TextBox" value="/txt/"/> <font color=red>*</font> 如 /Txt/
			<label><input type='checkbox' name='subfolder' id='subfolder' value='1' checked>包含子栏目</label> <input type='button' class='button' value='获取Txt文件列表' onclick='getAddress()'/>
			</td>
		 </tr>
		  <tr class='tdbg'>
			<td height="30" class='clefttitle' align='right'><strong>TXT文件列表：</strong></td>
            <td>
			<select multiple name="AddressUrls" id="AddressUrls" style="width:300px;height:200px">
			</select>
			<br/>
			<label><input type="checkbox" name="repeat" value="1" checked>标题重复不导入</label>
			</td>
		 </tr>
		  <tr class="tdbg">
			  <td colspan="2" style="text-align:center;height:40px;"><input type='submit' onclick="return(check());" value=' 开始 执行导入 ' class='button' name='button1'>
			  </td>
			 </tr>
			 </table>
			</form>
			
			
			<script>
			function getAddress(){
			  if ($('#subfolder[checked=true]').val()==undefined){
			   subfolder=0;
			  }else{
			   subfolder=1;
			  }
			  $(parent.document).find("#ajaxmsg").toggle();
			  $.ajax({
			  url: "ks.articleadd.asp",
			  cache: false,
			  data: 'action=geturls&folderpath='+$('#TxtDir').val()+'&subfolder='+subfolder,
			  success: function(d){
			   $(parent.document).find("#ajaxmsg").toggle();
			    if (d=='errorfolder'){
				   alert('对不起，您输入的目录不存在！');
				    $("#AddressUrls").empty();
				}else{
			      $("#AddressUrls").empty().append(d);
				  $("#AddressUrls option").attr("selected",true);
			     }
			  }})
			}
			</script>
			
			
			<div class="attention">
			<strong>说明：</strong><br/>1、请将您预先设计好的Txt文件放在网站的指定目录下，如Txt目录。
			<br/>2、如果有启用生成静态，批量添加后，请到后台底部的发布管理里重新生成一个内容页HTML操作。
			<br/>3、一个TXT文件如果要添加为多篇文章，请按格式录入好，格式如下：
			
			<br/><strong>如“1.txt” 文件格式：</strong>
			<pre>
			  第一篇文章标题
			  #####
			  第一篇文章内容
			  @@@@@
			  第二篇文章标题
			  #####
			  第二篇文章内容
			  @@@@@
			  第三篇文章标题
			  #####
			  第三篇文章内容
			</pre>
			<br/><strong>如“2.txt” 文件格式：</strong>
			<pre>
			  第一篇文章标题
			  #####
			  第一篇文章内容
			  @@@@@
			  第二篇文章标题
			  #####
			  第二篇文章内容
			  @@@@@
			  第三篇文章标题
			  #####
			  第三篇文章内容
			</pre>
			
			<font color=green>如上面的两个Txt文件放在 Txt目录下，则添加后将为添加为6篇文章。</font>
			
			</div><br/>
		
		  <%
		  
		  
	   end sub
	   
	   
	  sub innerjs(msg)
	   		Response.Write "<script>$('#message').html('" & msg & "');</script>" &vbcrlf
			Response.Flush

	  end sub
	  sub dotxtsave()
	      dim channelid:channelid=ks.chkclng(ks.s("channelid"))
		  dim classid:classid=ks.s("classid")
		  dim AddressUrls:AddressUrls=ks.s("AddressUrls")
		  dim inputer,repeat
		  inputer=KS.G("inputer")
		  repeat=KS.ChkClng(request("repeat"))
		  if channelid=0 then ks.die "<script>alert('请选择要发布的模型!');history.back();</script>"
		  if ks.isnul(classid) then ks.die "<script>alert('请选择要发布的栏目!');history.back();</script>"
          if ks.isnul(AddressUrls) then ks.die "<script>alert('没有找到可导入的Txt文件!');history.back();</script>"
		  dim addressArr:addressArr=split(AddressUrls,",")
		  %>
		  <div style="text-align:center">			 
			 <div style="margin-top:50px;border:1px dashed #cccccc;width:500px;height:80px">
			 <br>
			<div id="message">
			  <br>操作提示栏！
			</div>
			</div>
	    </div>
		<br/><br/><br/>
		  <%
		  dim i,j,total,n,errnum
		  total=ubound(addressArr)
		  n=0
		  errnum=0
		  for i=0 to total
		   
		   dim str:str=KS.ReadFromFile(addressArr(i))
		   if not ks.isnul(str) then
		      if instr(str,"#####")<>0 then  '有标题才导入
			     dim strarr:strarr=split(str,"@@@@@")
				 for j=0 to ubound(strarr)
				   dim title:title=replace(split(strarr(j)&"#####","#####")(0)&"",chr(10),"")
				   dim content:content=split(strarr(j)&"#####","#####")(1)
				  
				   if not ks.isnul(title) then
				     if repeat=1 then  '判断标题重复的记录不导入
					   if conn.execute("select top 1 id from " & KS.C_S(ChannelID,2) & " where tid='" & classid & "' and title='" & title & "'").eof then
					    call addrecord(channelid,classid,title,content,inputer)
					    n=n+1
					   else
					     errnum=errnum+1
					   end if
					 else
				       call addrecord(channelid,classid,title,content,inputer)
					   n=n+1
					 end if
				   end if
				 next
			  end if
		   end if
		   
		   call innerjs("共有<font color=green>" & total+1 & "</font> 个Txt文件，正在导入第<font color=red>" & i+1 & "</font>个Txt文件,文件路径：" & addressArr(i) & "!<br/>")
		  next
		  
		  call innerjs("执行完毕，共导入了 <font color=green>" & total+1 & "</font> 个Txt文件,累计成功导入<font color=red>" & n & "</font>篇文章，<font color=red>" & errnum &"</font>篇文章标题重复没有导入!<br/><br/><input type=button value="" 返 回 "" onclick=""history.back()"" class=""button""/>")

		  
		  
	   end sub
	   
	   
	   
	   '
	   Sub add()
			 response.write "<form action=""?Action=dosave"" method=""post"" name=""DownParamForm"">"
			addHead
			%>
			<tr class='tdbg'>
			<td height="30" class='clefttitle' align='right'><strong>添加篇数</strong></td>
            <td><select name="num" onchange="getAdd(this.value);">
			  <%
			   dim nn:nn=0
			   for nn=1 to 300
			    if nn=10 then
			    response.write "<option value=""" & nn & """ selected>" & nn & " 篇</option>"
				else
			    response.write "<option value=""" & nn & """>" & nn & " 篇</option>"
				end if
			   next
			  %>
			</select> &nbsp;<strong>录入员：</strong><input type="text" name="inputer" value="<%=KS.C("AdminName")%>"/></td>
			 </tr>
			 
			 <tbody id="showadd">
			 </tbody>
			 <tr class="tdbg">
			  <td colspan="2" style="text-align:center;height:40px;"><input type='submit' onclick="return(check());" value=' 确定保存 ' class='button' name='button1'>
			  </td>
			 </tr>
			 </table>
			</form>
			<div class="attention">
			<strong>说明：</strong><br/>1、文章标题必须输入，没有输入文章标题的文章不入库。
			<br/>2、如果有启用生成静态，批量添加后，请到后台底部的发布管理里重新生成一个内容页HTML操作。
			</div><br/><br/>
			<%
		End Sub
		
		
		Sub dosave()
		  dim channelid:channelid=ks.chkclng(ks.s("channelid"))
		  dim classid:classid=ks.s("classid")
		  dim num:num=ks.chkclng(ks.s("num"))
		  dim n,inputer,sucnum,fname,title,content,intro
		  inputer=KS.G("inputer")
		  sucnum=0
		  if channelid=0 then ks.die "<script>alert('请选择要发布的模型');history.back();</script>"
		  if ks.isnul(classid) then ks.die "<script>alert('请选择要发布的栏目!');history.back();</script>"
		  for n=1 to num
		     if not ks.isnul(request("title"&n)) then
			   sucnum=sucnum+1
			   fname=KS.GetFileName(KS.C_C(classid,24), Now, KS.C_C(classid,23))
			   title=request("title"&n)&""
			   content=request("content"&n) &""
			   call addrecord(channelid,classid,title,content,inputer)
			 end if
		  next
		  
		  ks.die "<script>alert('恭喜，共成功添加了 " & sucnum & " 篇文章!');location.href='KS.ArticleAdd.asp';</script>"
		 
		End Sub
		
		sub addrecord(channelid,classid,title,content,inputer)
		    dim intro:intro=ks.gottopic(content,200)
			
		    dim fname:fname=KS.GetFileName(KS.C_C(classid,24), Now, KS.C_C(classid,23))
		    dim rs:set rs=server.CreateObject("adodb.recordset")
            rs.open "select top 1 * from " & KS.C_S(ChannelID,2) & " Where 1=0",conn,1,3
			    rs.addnew
				rs("title")=title
				rs("articlecontent")=replace(content,chr(10),"<br/>")
				rs("intro")=intro
				rs("inputer")=inputer
				rs("tid")=classid
				rs("verific")=1
				rs("deltf")=0
				rs("PostTable") = LFCls.GetCommentTable()
				rs("verific")=1
				rs("templateid")=KS.C_C(ClassID,5)
				rs("waptemplateid")=KS.C_C(ClassID,22)
				rs("fname")=fname
				rs("adddate")=now
				RS("ModifyDate")=now
				rs("rank")="★★★"
				rs("hits")=0
				rs("hitsbyday")=0
				rs("hitsbyweek")=0
				rs("hitsbymonth")=0
				rs("recommend")=0
				rs("rolls")=0
				rs("strip")=0
				rs("popular")=0
				rs("slide")=0
				rs("istop")=0
				rs("comment")=1
				rs("OrderID")        = KS.ChkClng(Conn.Execute("Select Max(OrderID) From " & KS.C_S(ChannelID,2) & " Where Tid='" & classid &"'")(0))+1
			   rs.update
			   rs.movelast
			   If Left(Ucase(Fname),2)="ID"  Then
					   RS("Fname") = RS("ID") & KS.C_C(classid,23)
					   RS.Update
				End If
			    Call LFCls.AddItemInfo(ChannelID,RS("ID"),Title,classid,Intro,"","",now,KS.C("AdminName"),0,0,0,0,0,0,0,0,0,0,0,1,RS("Fname"))
			   rs.close
		end sub


End Class
%> 

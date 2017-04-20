<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
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
Set KSCls = New Install
KSCls.Kesion()
Set KSCls = Nothing

Class Install
        Private ChannelID,ModelTable,Param,XML,Node,StartTime,IsCheckEnvironmentPass
		Private CurrPage,MaxPerPage,TotalPut,PageNum,Key,stype,OrderStr
		Private Sub Class_Initialize()
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		End Sub
		
	 Function InitialObject(str)
		'iis5创建对象方法Server.CreateObject(ObjectName);
		'iis6创建对象方法CreateObject(ObjectName);
		'默认为iis6，如果在iis5中使用，需要改为Server.CreateObject(str);
		Set InitialObject=CreateObject(str)
	 End Function
	 
	'**************************************************
	'函数：FoundInArr
	'作  用：检查一个数组中所有元素是否包含指定字符串
	'参  数：strArr     ----字符串
	'        strToFind    ----要查找的字符串
	'       strSplit    ----数组的分隔符
	'返回值：True,False
	'**************************************************
	Public Function FoundInArr(strArr, strToFind, strSplit)
		Dim arrTemp, i
		FoundInArr = False
		If InStr(strArr, strSplit) > 0 Then
			arrTemp = Split(strArr, strSplit)
			For i = 0 To UBound(arrTemp)
			If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
				FoundInArr = True:Exit For
			End If
			Next
		Else
			If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then FoundInArr = True
		End If
	End Function
	
	'**************************************************
	'函数名：R
	'作  用：过滤非法的SQL字符
	'参  数：strChar-----要过滤的字符
	'返回值：过滤后的字符
	'**************************************************
	Public Function R(strChar)
		If strChar = "" Or IsNull(strChar) Then R = "":Exit Function
		Dim strBadChar, arrBadChar, tempChar, I
		'strBadChar = "$,#,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
		strBadChar = "+,',--,%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
		arrBadChar = Split(strBadChar, ",")
		tempChar = strChar
		For I = 0 To UBound(arrBadChar)
			tempChar = Replace(tempChar, arrBadChar(I), "")
		Next
		tempChar = Replace(tempChar, "@@", "@")
		R = tempChar
	End Function
	 
	 
	'取得Request.Querystring 或 Request.Form 的值
	Public Function G(Str)
	 G = Replace(Replace(Replace(Replace(Request(Str), "'", ""), """", ""),"%","％"),"*","＊")
	End Function
	 
	'**************************************************
	'函数名：GetAutoDoMain()
	'作  用：取得当前服务器IP 如：http://127.0.0.1
	'参  数：无
	'**************************************************
	Public Function GetAutoDomain()
		Dim TempPath
		If Request.ServerVariables("SERVER_PORT") = "80" Then
			GetAutoDomain = Request.ServerVariables("SERVER_NAME")
		Else
			GetAutoDomain = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
		End If
		 If Instr(UCASE(GetAutoDomain),"/W3SVC")<>0 Then
			   GetAutoDomain=Left(GetAutoDomain,Instr(GetAutoDomain,"/W3SVC"))
		 End If
		 GetAutoDomain = "http://" & GetAutoDomain
	End Function
	
		
		Public Sub Kesion()
		
		
		%>
		<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
        <html xmlns="http://www.w3.org/1999/xhtml" >
        <head><title>
            KesionCMS (ASP版) X2.0 系列产品安装向导
        </title><link href="images/guide.css" rel="stylesheet" />
        <script src="../ks_inc/jquery.js" type="text/javascript"></script>
		<script src="../ks_inc/common.js" type="text/javascript"></script>
        </head>
        <body>  
        <form name="form" method="post" action="index.asp" id="form">
        <div class="guide">
         <div class="guidetitle">
                    <div class="l"></div><div class="r">当前安装版本：KesionCMS X2.0 官方版本:<script src="http://www.kesion.com/websystem/GetofficialInfo.asp?action=getverbyscript"></script></div>
                </div>
                <div class="guidesm"><div class="l">欢迎使用KesionCMS X2.0安装向导。本向导即将安装软件到您的系统中，安装前仔细阅读一下协议，点击同意进入下一步。</div>
                <div class="r"><img src="images/download.png" width="80"></div>
                <div class="clear"></div>
              </div>
          <div class="clear"></div>
		<%
		If ISFromFile("install.lock") ="ok" Then 
			%>
             <script>
				KesionJS.Alert("安装程序已运行过了，如果要重新安装，请先删除install/install.lock文件！", "location.href = '/';");
            </script>
			<%
			Response.end()
		end if
		
		 select case Request("action")
		 Case "s1"
		  	Call box_1()
		 Case "s2"
		  	Call box_2()	
		 Case "s3"
		  	Call box_3()	
		 Case "s4"
		  	Call box_4()
		 Case "s5"
		  	Call box_5()		
		 case else
		 	Call index_box()
		 end select
		%>
        </form>
		
         </div>
		</body>
        </html>
		<%


	   End Sub
	   
	   Sub index_box()
	   		%>
			<input type="hidden" name="action" value="s1"  />
                
                <div id="Step1">
                <div class="datiao">
                    <ul>
                        <li class="curr">阅读许可协议</li>
                        <li>检查安装环境</li>
                        <li>创建数据库</li>
                        <li>网站参数配置</li>
                        <li>完成安装</li>
                    </ul>
                </div>
                <div class="clear"></div>
                <div class="guideboxbig">
                    <div class="guidebox">
                        <h4>KesionCMS X2.0安装许可协议</h4>
                        <strong>版权所有（C） 2006-<%=year(now)%> 厦门科汛软件有限公司</strong>
        <br>KesionCMS 系列产品是厦门科汛软件有限公司独立开发，依法独立拥有KesionCMS 系列产品的所有着作权。
        <br>KesionCMS 系列产品的着作权已在中华人民共和国国家版权局注册，软件制作权登记号：<span style='color:green'>2016SR000393</span>。着作权受到法律和国际公约保护。使用者：无论个或组织、盈利与否、用途如何（包括以学习和研究为目的），均需仔细阅读本许可协议，在理解、同意、并遵守本许可协议的全部条件和条款后，方可开始使用KesionCMS 系列产品。
        <br>有关本软件的用户许可协议、商业授权与技术服务的详细内容，均由厦门科汛软件有限公司独家提供。厦门科汛软件有限公司拥有在不事先通知的情况下，修改许可协议和服务价目表的权力，修改后的协议或价目表对自改变之日起的新授权用户生效。 
        <br>电子文本形式的许可协议如同双方书面签署的协议一样，具有完全的和等同的法律效力。您一旦开始确认本协议并安装、使用、修改或分发本软件（或任何基于本软件的衍生着作），则表示您已经完全接受本许可协议的所有的条件和条款。如果您有任何违反本许可协议的行为，厦门科汛软件有限公司有权收回对您的许可授权，责令停止损害，并追究您的相关法律及经济责任。
        <br>
        <br><strong>1、许可</strong>
        <br>1.1	本软件仅供给个人用户非商业使用。如果您是个人用户，那么您可以在完全遵守本用户许可协议的基础上，将本软件应用于非商业用途，而不必支付软件授权许可费用。 
        <br>1.2	您可以在本协议规定的约束和限制范围内修改本软件的源代码和界面风格以适应您的网站要求。
        <br>1.3	您可以在本协议规定的约束和限制范围内通过任何的媒介和渠道复制与分发本软件的源代码的副本（要求是逐字拷贝的副本）。
        <br>1.4	您拥有使用本软件构建的网站全部内容所有权，并独立承担与这些内容的相关法律义务。
        <br>1.5	在获得商业授权之后，您可以将本软件应用于商业用途。 
        <br>
        <br><strong>2、约束和限制</strong>
        <br>2.1	未获商业授权之前，不得将本软件用于商业用途，不得用于任何非个人所有的项目之中，例如属于企业、政府单位所有的网站。
        <br>2.2	未获商业授权之前，不得以任何形式提供与本软件相关的收费服务，包括但不限于以下行为：为用户提供本软件的相关咨询或培训服务并收费一定费用；用本软件为他人建站并收取一定费用；用本软件提供SaaS（软件做为服务）服务。
        <br>2.3	不得对本软件或与之关联的商业授权进行出租、出售、抵押或发放子许可证。
        <br>2.4	禁止任何以获利为目的的分发本软件的行为。
        <br>2.5	禁止在本软件的整体或任何部分基础上以发展任何派生版本、修改版本或第三方版本用于重新分发。
        <br>
        <br><strong>3、无担保及免责声明</strong>
        <br>3.1	用户出于自愿而使用本软件，您必须了解使用本软件的风险，且同意自己承担使用本软件的风险。
        <br>3.2	用户利用本软件构建的网站的任何信息内容以及导致的任何版权纠纷和法律争议及后果与厦门科汛软件有限公司无关，厦门科汛软件有限公司对此不承担任何责任。
        <br>3.3	在适用法律允许的最大范围内，厦门科汛软件有限公司在任何情况下不就因使用或不能使用本软件所发生的特殊的、意外的、非直接或间接的损失承担赔偿责任（包括但不限于，资料损失，资料执行不精确，或应由您或第三人承担的损失，或本程序无法与其他程序运作等）。即使用户已事先被厦门科汛软件有限公司告知该损害发生的可能性。
        <br><div style="text-align:right">福建厦门科汛软件有限公司</div>
        <br><div style="text-align:right"><%=formatdatetime(now,2)%></div>
        <br><br>
                    </div>
                    <div class="clear blank10"></div>
                    <div>
                    <input name="BtnAgree" value="我同意" id="BtnAgree" class="btnbg" type="submit">
                    <input name="" value="我不同意" onClick="window.close()" class="btnbg" type="button">
                    </div>
                </div>
                
        </div>
			<%
	   End Sub	
	   
	   
	   Sub box_1()
	   IsCheckEnvironmentPass=true
	   %>
       <input type="hidden" name="action" value="s2"  />
	   <div id="Step2">
                <div class="datiao">
                    <ul>
                        <li>阅读许可协议</li>
                        <li class="curr">检查安装环境</li>
                        <li>创建数据库</li>
                        <li>网站参数配置</li>
                        <li>完成安装</li>
                    </ul>
                </div>
                <div class="clear"></div>
                <div class="hjlist">
                    <h5>环境检查</h5>
                    <ul class="tit">
                        <li class="li01">项目</li>
                        <li class="li02">KESION所需配置</li>
                        <li class="li04">当前服务器</li>
                    </ul>
                    <ul>
                        <li class="li01">ASP 版本</li>
                        <li class="li02">ADO 数据对象</li>
                        <li class="li04">
                          <%
						 On Error Resume Next
						InitialObject("adodb.connection")
						if err=0 then 
						  %><img src="images/v.png" align="absmiddle">adodb.connection<%
						else
						 %><img src="images/no.gif" align="absmiddle">不支持<%
						end if	 
						err=0
					  %>  
                        </li>
                    </ul>
                    <ul>
                        <li class="li01">FSO</li>
                        <li class="li02">FSO文本文件读写</li>
                        <li class="li04">
                         <%
						 On Error Resume Next
						InitialObject("Scripting.FileSystemObject")
						if err=0 then 
						  %><img src="images/v.png" align="absmiddle"><% Response.Write("Scripting.FileSystemObject")
						else
						 %><img src="images/no.gif" align="absmiddle">不支持<%
						end if	 
						err=0
					 	 %>  
                        
                        </li>
                    </ul>
                    
                    <div class="clear"></div>
                    <h5>目录、文件权限检查</h5>
                    <ul class="tit">
                        <li class="li05">目录文件</li>
                        <li class="li06">所需状态</li>
                        <li class="li07">当前状态</li>
                    </ul>
                    <ul>
                        <li class="li05">/Admin/</li>
                        <li class="li06"><img src="images/v.png" align="absmiddle">可写</li>
                        <li class="li07">
                        	<% Response.Write isFsowrite("../Admin/1.txt")%>
                        </li>
                    </ul>
                    <ul>
                        <li class="li05">/Uploadfiles/</li>
                        <li class="li06"><img src="images/v.png" align="absmiddle">可写</li>
                        <li class="li07">
                        <% Response.Write isFsowrite("../Uploadfiles/1.txt")%>
                        </li>
                    </ul>
                    <ul>
                        <li class="li05">/template/</li>
                        <li class="li06"><img src="images/v.png" align="absmiddle">可写</li>
                        <li class="li07">
                        <% Response.Write isFsowrite("../template/1.txt")%>
                        </li>
                    </ul>
                    <ul>
                        <li class="li05">/Config/</li>
                        <li class="li06"><img src="images/v.png" align="absmiddle">可写</li>
                        <li class="li07">
                        <% Response.Write isFsowrite("../Config/1.txt")%>
                        </li>
                    </ul>
                    <ul>
                        <li class="li05">/API/</li>
                        <li class="li06"><img src="images/v.png" align="absmiddle">可写</li>
                        <li class="li07">
                        <% Response.Write isFsowrite("../API/1.txt")%>
                        </li>
                    </ul>
                    <ul>
                        <li class="li05">/KS_Data/</li>
                        <li class="li06"><img src="images/v.png" align="absmiddle">可写</li>
                        <li class="li07">
                        <% Response.Write isFsowrite("../KS_Data/1.txt")%>
                        </li>
                    </ul>
                  <%If CheckDir("../html/") then%>
                    <ul>
                        <li class="li05">/HTML/</li>
                        <li class="li06"><img src="images/v.png" align="absmiddle">可写</li>
                        <li class="li07">
                        <% Response.Write isFsowrite("../HTML/1.txt")%>
                        </li>
						
                    </ul>
					<%end if%>
					<%If CheckDir("../mnkc/") then%>
                    <ul>
                        <li class="li05">/mnkc/</li>
                        <li class="li06"><img src="images/v.png" align="absmiddle">可写</li>
                        <li class="li07">
                        <% Response.Write isFsowrite("../mnkc/1.txt")%>
                        </li>
                    </ul>
					<%end if%>
                    <ul>
                        <li class="li05">/JS/</li>
                        <li class="li06"><img src="images/v.png" align="absmiddle">可写</li>
                        <li class="li07">
                        <% Response.Write isFsowrite("../JS/1.txt")%>
                        </li>
                    </ul>
                    <ul>
                        <li class="li05">/conn.asp</li>
                        <li class="li06"><img src="images/v.png" align="absmiddle">可写</li>
                        <li class="li07">
                        <% 
						On Error Resume Next
							call setName("../conn.asp","conn2.asp",0)
						if err=0 then
							Response.Write "<img src=""images/v.png"" align=""absmiddle"">可写"
						else
							Response.Write "<img src=""images/no.gif"" align=""absmiddle"">不可写"
							IsCheckEnvironmentPass=false
						end if	
							call setName("../conn2.asp","conn.asp",0) 
						err=0
						%>
                        </li>
                    </ul>
                    
                    <div class="clear blank10"></div>
                    <div style="padding:5px">
                    <input name="" value="上一步" onClick="history.back()" class="btnbg" type="button">
                    <input name="BtnStep2Next" <%if IsCheckEnvironmentPass=false then%> onclick="KesionJS.Alert('对不起，您的环境检测没有通过，请检查！','');return false;"<%end if%>value="下一步" id="BtnStep2Next" class="btnbg" type="submit">
                    </div>
                </div>
                
        </div>
	   <%	
	   End Sub
	   
	   Sub box_2()	   
	   %>
        <input type="hidden" name="action" value="s3"  />
       <div id="Step3">
             <div class="datiao">
                <ul>
                    <li>阅读许可协议</li>
                    <li>检查安装环境</li>
                    <li class="curr">创建数据库</li>
                    <li>网站参数配置</li>
                    <li>完成安装</li>
                </ul>
             </div>
            <div class="clear"></div>
            <div class="sjlist">
                <h5>填写数据库信息</h5>
                <ul>
                    <li><span>数据库类型选择：</span>
                     <input type="radio" name="DBlx" value="0" checked="checked" onClick="$('#Access_s').show();$('#SQL_s').hide();" /> Access 数据库  <input type="radio" name="DBlx" value="1" onClick="$('#Access_s').hide();$('#SQL_s').show();"  disabled="disabled"/> SQL 数据库 
                    </li>
                </ul>
                <ul id="Access_s" >
                    <li><span>数据库文件名：</span>/KS_Data/<input name="TxtDBName_a" value="KesionCMSx20.mdb" id="TxtDBName_a" class="text" type="text">
                    如：KesionCMS.asa,尽量不要以MDB为扩展名。
                    </li>
                    
                    </li>
                </ul>
                <ul id="SQL_s" style="display:none;">
                 <li><span>数据库版本：</span>
				  &nbsp;&nbsp;&nbsp;<select name="sqlversion">
				   <option value="0">SQL 2000</option>
				   <option value="1">SQL 2005/2008及以上版本</option>
				  </select> 请务必选择正确的数据库版本，否则可能导致安装不成功！
				 </li>
                 <li><span>数据库服务器：</span><input name="TxtDBService" value="(local)" id="TxtDBService" class="text" type="text">数据库服务器地址, 一般为 (local)</li>
                    <li><span>数据库名：</span><input name="TxtDBName" value="kesioncmsX20" id="TxtDBName" class="text" type="text">
                    请确保数据库已存在，否则请先创建数据库。
                    </li>
                    <li><span>数据库用户名：</span><input name="TxtDBUser" value="sa" id="TxtDBUser" class="text" type="text"></li>
                    <li><span>数据库密码：</span><input name="TxtDBPass" value="989066" id="TxtDBPass" class="text" type="text"></li>
                    </li>
                </ul>  
				<ul>
                    <li><span>初始数据：</span>&nbsp;<strong> <input id="CkbData" name="CkbData" checked="checked" value="1" type="checkbox"><label for="CkbData">安装体验数据包(推荐)</label></strong>

				</ul>
                <div class="clear"></div>
                
                <div style="padding:5px">
                <input name="BtnStep3Next" value="下一步" onclick="doAjax();" id="BtnStep3Next" class="btnbg" type="button" />
                </div>
                <script>
                    function doAjax() {
					      $("#form").attr("target","hidframe");

					    if ($("input[name='DBlx']:checked").val()==1){
							if (jQuery("#TxtDBService").val() == '') {
								KesionJS.Alert("请输入数据库服务器地址！", "jQuery('#TxtDBService').focus();");
								return false;
							}
							if (jQuery("#TxtDBName").val() == '') {
								KesionJS.Alert("请输入数据库名称！", "jQuery('#TxtDBName').focus();");
								return false;
							}
							if (jQuery("#TxtDBUser").val() == '') {
								KesionJS.Alert("请输入数据库用户名！", "jQuery('#TxtDBUser').focus();");
								return false;
							}
							if (jQuery("#TxtDBPass").val() == '') {
								KesionJS.Alert("请输入数据库密码！", "jQuery('#TxtDBPass').focus();");
								return false;
							}
                        } else{
						
						   if (jQuery("#TxtDBName_a").val() == '') {
								KesionJS.Alert("请输入数据库名称！", "jQuery('#TxtDBName_a').focus();");
								return false;
							}
						}
	                    $("#form").submit();
						
                       // jQuery("#BtnStep3Next").attr("disabled", true);
						
                     }
                    
                </script>                       
            </div>
            
            
    </div>
	
	<iframe name="hidframe" id="hidframe" src="about:blank" style="display:none;width:400px;height:300px;"></iframe>
	<div id="showtips" style="display:none;padding:2px;left:300px;background:#fff;width:450px;height:250px;top:180px;border:3px solid #ccc;position:absolute;z-index:1000;overflow-x:hidden;overflow-y:scroll">
	  <h3><span style="float:right;cursor:pointer;" onclick="$('#showtips').slideUp('fast')"><img width='12' src="images/no.gif" align="absmiddle"/>关闭</span> <img src="images/01.png" align="absmiddle" width="20"/>安装提示信息:</h3>
	  <div id="UpdateInfo"></div>
	  <div id="msg_end" style="height:0px; overflow:hidden"></div>
	</div>
       
	   <%
	   End Sub
	   
	    '数据转换操作
		Sub TransferData(Table3)
			       on error resume next
				 	dim ArrStr,k,rsNew
					Set rsNew = Server.CreateObject("ADODB.Recordset")
					rsNew.Open "Select * From " & Table3,Conn_old,1,3
					For k=0 To rsNew.Fields.count-1
							If k=0 Then
							 ArrStr="[" & rsNew.Fields(k).Name& "]"
							Else
							 ArrStr=ArrStr &",["&rsNew.Fields(k).Name & "]"
							End If
						  Next
				    rsNew.close
					CONN.CommandTimeout = 600 
					CONN.execute("SET IDENTITY_INSERT [" & Table3 & "] ON")
					CONN.execute("INSERT INTO [" & Table3 & "] (" & ArrStr & ") " & _
						"SELECT " & ArrStr & " " & _
						"FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0','Data Source=""" & server.mappath("../ks_data/KesionCMSx20.mdb") & """')...[" & Table3 & "]")
					if err<>0 then 
					 %>
						<script>
							KesionJS.Alert("数据表<%=Table3%>导入失败，原因：<%=err.description%>,建议手工安装SQL数据库!","");
                        </script>
				    <%
					 err.clear
					 response.End()
					end if
					 Call InnerHtml("<font color=green>表[" & Table3 & "]的数据安装成功!</font>")
					response.flush()
			End Sub
	   
			'数据转换操作
			Sub TransferNotData(Table3)
				 Dim RS2:Set RS2=Server.CreateObject("ADODB.RECORDSET")
				 Dim RS3:Set RS3=Server.CreateObject("ADODB.RECORDSET")

				 RS2.Open "Select * From " & Table3,Conn_Old,1,1
				 RS3.Open "Select * From " & Table3,Conn,3,3
				  Dim I,start,ArrStr
				   For I=0 To RS3.Fields.count-1
				    If I=0 Then
					 ArrStr=rs3.Fields(i).Name
					Else
				     ArrStr=ArrStr &","&rs3.Fields(i).Name
					End If
				  Next
				   
			      Do While Not RS2.Eof
					   RS3.AddNew
						   For I=0 To RS2.fields.count-1
						    If instr(ArrStr,rs3.Fields(i).Name) Then
						     RS3(rs2.Fields(i).Name) = rs2.Fields(i).value

							End If
						   Next
					  RS3.Update
					   RS2.MoveNext
					 Loop

				 RS3.Close:Set RS3=Nothing
				 RS2.Close:Set RS2=Nothing
				 Call InnerHtml("<font color=green>表[" & Table3 & "]的数据安装成功!</font>")
				 
			End Sub	
	  
		Sub InnerHtml(msg)
			Response.Write "<SCRIPT>$('#UpdateInfo',parent.document).html($('#UpdateInfo',parent.document).html()+""<li>"&msg&"</li>""); $('#msg_end',parent.document)[0].scrollIntoView(); </SCRIPT>"
			Response.Flush
		End Sub


	   Sub box_3()
	   if Request("DBlx")="0" then ' Access
		   if Request("TxtDBName_a")<>"" then 
		   		if fileExists("../KS_Data/"&Request("TxtDBName_a"))<>true then
					On Error Resume Next
					Call Copy("../KS_Data/KesionCMSx20.mdb","/KS_Data/"&Request("TxtDBName_a"),0)
					if err<>0 then
						%>
						<script>
							KesionJS.Alert("Access 数据库创建失败,请手动打开/KS_Data/数据库文件改名!","");
                        </script>
						<%
					end if
					Call FileDel("../KS_Data/KesionCMSx20.mdb",0)
					if err<>0 then 
						%>
						<script>
							KesionJS.Alert("/KS_Data/KesionCMSx20.mdb 删除失败,请手动删除文件!","");
                        </script>
						<%
					end if
					
				end if
		   end if
	   else 'SQL数据库
	     Dim Conn_Str: Conn_Str="Provider = Sqloledb; User ID = " & Request("TxtDBUser") & "; Password = " & Request("TxtDBPass") & "; Initial Catalog = " & Request("TxtDBName") & "; Data Source = " & Request("TxtDBService") & ";"
	     Call OpenConn_install(Conn_Str,1)
         
		 response.write "<script>$('#showtips',parent.document).show();</script>"
         Call InnerHtml("<font color=blue>正在安装SQL数据库架构...</font>")

		 Dim SQLText:SQLText=ReadFromFile("data.sql")
         If SQLText<>"" Then
		   Dim ii,SQLArr:SQLArr=Split(SQLText,"GO")
		   For II=0 To Ubound(SQLArr)
		    on error resume next
		    Conn.Execute(SQLArr(ii))
			if err then 
			    Call InnerHtml("在线安装SQL数据库失败，原因：<font color=red>" & err.description & "</font>,建议手工安装!")
				response.End()
			end if
			
		   Next
		 End If
		 Call InnerHtml("<font color=green>数据库架构安装成功!</font>")
		
		 
		 '============================转移数据==================================
		 
		  '===============SQL2005版本支持=========================================
			if Request("sqlversion")="1" then 'SQL2005版本支持
			    on error resume next
				conn.execute("exec sp_configure 'show advanced options',1")
				conn.execute("reconfigure")
				conn.execute("exec sp_configure 'Ad Hoc Distributed Queries',1")
				conn.execute("reconfigure")
				if err then
		         response.write "<script>$('#showtips',parent.document).hide();</script>"
			     Call InnerHtml("您的数据库环境不支持在线安装,建议手工安装!")
				 response.end
				end if
			end if
			'=================================================================================
		 
		   Call InnerHtml("<font color=blue>正在安装表数据...</font>")
			OpenOldConn()
			if Request("CkbData")<>"1" then '删除初始数据
			 DelDefaultData(conn_old)
		    end if 
			
			
			   dim rs:Set rs = Conn_old.OpenSchema(4)
			   dim tablename:tablename=""
			   dim tablearr,temptable
				Do Until rs.EOF
					temptable=rs("Table_name")
					if temptable <> tablename and lcase(temptable)<>"ks_notdown" and lcase(left(temptable,3))="ks_"  then
					    tablearr=tablearr & temptable & ","
						Tablename = temptable
					end if
				rs.MoveNext
				Loop
				rs.close:set rs=nothing

				
				tablearr=left(tablearr,len(tablearr)-1)
				dim k,NotStr,i
				NotStr="KS_PaymentPlat,KS_Channel,KS_Online,KS_Label,KS_LabelFolder,KS_Config,KS_JSFile,KS_MovieParam,KS_Origin,KS_DownSer,KS_AskClass,KS_ItemInfoR,KS_WapTemplate"
				tablearr=split(tablearr,",")
				for i=lbound(tablearr) to ubound(tablearr)
				    If FoundInArr(NotStr, tablearr(i),",") = False Then
						for k=0 to ubound(tablearr)
						  If FoundInArr(NotStr, tablearr(k),",") = False Then
							CONN.execute("SET IDENTITY_INSERT [" & tablearr(k) & "] Off")
						  end if
						next 
						Call TransferData(tablearr(i))
				    end if
			    next
				
				
				
				tableArr=split(NotStr,",")
				for i=lbound(tablearr) to ubound(tablearr)
						Call TransferNotData(tablearr(i))
			    next
				 Call InnerHtml("<font color=green>所有表数据迁移成功！</font>")
			'========================================================================
				 Call InnerHtml("<font color=green>数据库安装完毕！</font>")
	     
	   end if
	      closeconn
		  
		  if Request("DBlx")="0" then ' Access
	       response.write "<script>top.location.href='index.asp?action=s4&DBlx=" & request("DBlx") & "&CkbData=" & server.urlencode(request("CkbData"))&"&TxtDBName_a=" & server.urlencode(request("TxtDBName_a")) &"&TxtDBService=" & server.urlencode(request("TxtDBService"))&"&TxtDBName=" & server.urlencode(request("TxtDBName"))&"&TxtDBUser=" & server.urlencode(request("TxtDBUser"))&"&TxtDBPass=" & request("TxtDBPass") &"';</script>"
		  Else
	       response.write "<script>KesionJS.Alert('恭喜，数据库安装成功，点击确定进入下一步！',""top.location.href='index.asp?action=s4&DBlx=" & request("DBlx") & "&CkbData=" & server.urlencode(request("CkbData"))&"&TxtDBName_a=" & server.urlencode(request("TxtDBName_a")) &"&TxtDBService=" & server.urlencode(request("TxtDBService"))&"&TxtDBName=" & server.urlencode(request("TxtDBName"))&"&TxtDBUser=" & server.urlencode(request("TxtDBUser"))&"&TxtDBPass=" & request("TxtDBPass") &"';"");</script>"
		 End If
	   
	   
	  end sub
	  
	  sub box_4()
	   
	   dim InstallDir
	   dim strDir,strAdminDir
	   strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
	   strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
	   InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
			
		If Instr(UCASE(InstallDir),"/W3SVC")<>0 Then
			   InstallDir=Left(InstallDir,Instr(InstallDir,"/W3SVC"))
		End If
	   %>
       <input type="hidden" name="action" value="s5"  />
       <input type="hidden" name="DBlx" value="<%=Request("DBlx")%>"  />
       <input type="hidden" name="CkbData" value="<%=Request("CkbData")%>"  />
       
       <input type="hidden" name="TxtDBName_a" value="<%=Request("TxtDBName_a")%>"  />
       <input name="TxtDBService" value="<%=Request("TxtDBService")%>" id="TxtDBService" class="text" type="hidden"  />
       <input name="TxtDBName" value="<%=Request("TxtDBName")%>" id="TxtDBName" class="text" type="hidden" />
       <input name="TxtDBUser" value="<%=Request("TxtDBUser")%>" id="TxtDBUser" class="text" type="hidden" />
       <input name="TxtDBPass" value="<%=Request("TxtDBPass")%>" id="TxtDBPass" class="text" type="hidden"  />
      
	   <div id="Step4">
	
		
		 <div class="datiao">
			<ul>
				<li>阅读许可协议</li>
				<li>检查安装环境</li>
				<li>创建数据库</li>
				<li class="curr">网站参数配置</li>
				<li>完成安装</li>
			</ul>
		 </div>
		 <div class="clear"></div>
		 <div class="sjlist">
			<h5>网站参数配置</h5>
			<ul>
				<li><span>网站名称：</span><input name="TxtSiteName" value="科兴网络开发" id="TxtSiteName" class="text" type="text"><font color="red">*</font> 如：Kesion官方站</li>
				<li><span>网站域名：</span><input name="TxtSiteUrl" value="<%=GetAutoDomain %>" id="TxtSiteUrl" class="text" type="text"><font color="red">*</font> 后面不要带“/”。 
				如http://www.kesion.com。
				</li>
				<li><span>安装目录：</span><input name="TxtInstallDir" value="<%=InstallDir%>" id="TxtInstallDir" class="text" type="text"><font color="red">*</font> 后面不要带“/”。 
				系统会自动获取，建议不要修改。
				</li>
				<li><span>授 权 码：</span><input name="TxtSiteKey" value="0" id="TxtSiteKey" class="text" type="text">
				免费版本用户请留空或填“0”。
				</li>
				<li><span>后台目录：</span><input name="TxtManageDir" value="Admin/" id="TxtManageDir" class="text" type="text"><font color="red">*</font> 如：Manage,Admin，后面必须带"/"符号。</li>
                <li><span> 后台登录验证码：</span>
                 <input type="radio" name="isCode_a" value="True"  /> 启用  
                 <input type="radio" value="False"  name="isCode_a" checked="checked"/> 不启用
                </li>
               
				<li><span>管理认证码：</span>
                 <input type="radio" name="isCode" value="True" onclick="$('#rzm').show()"/> 启用  <input onclick="$('#rzm').hide()" type="radio" value="False"  name="isCode" checked="checked"   /> 不启用 
                <font id="rzm" style="display:none">认证码：<input name="TxtManageCode" value="8888"  id="TxtManageCode" class="text" style="width:100px;" type="text"></font></li>
			</ul>
			<div class="clear"></div>
			<h5>填写管理员信息</h5>
			<ul>
				<li><span>管理员账号：</span><input name="TxtUserName" value="admin"  id="TxtUserName" class="text" type="text"><font color="red">*</font> </li>
				<li><span>管理员密码：</span><input name="TxtUserPass" value="admin888" id="TxtUserPass" class="text" type="text"><font color="red">*</font> 管理员密码不能为空</li>
				<li><span>重复密码：</span><input name="TxtReUserPass" value="admin888" id="TxtReUserPass" class="text" type="text"></li>
			</ul>
			<div class="clear blank10"></div>
			
			<div style="padding:5px">
			<input name="Button1" value="下一步" onClick="return(doCheck());" id="Button1" class="btnbg" type="submit">
			</div>
		</div>
		
		<script>
			    function doCheck() {
			        if (jQuery("#TxtSiteName").val() == '') {
			            KesionJS.Alert("请输入网站名称！", "jQuery('#TxtSiteName').focus();");
			            return false;
			        }
			        if (jQuery("#TxtSiteUrl").val() == '') {
			            KesionJS.Alert("请输入网站网址！", "jQuery('#TxtSiteUrl').focus();");
			            return false;
			        }

					 if (jQuery("#TxtManageDir").val() == '') {
			            KesionJS.Alert("请输入后台目录！", "jQuery('#TxtManageDir').focus();");
			            return false;
			        }

			        if (jQuery("#TxtUserName").val() == '') {
			            KesionJS.Alert("请输入管理员登录账号！", "jQuery('#TxtUserName').focus();");
			            return false;
			        }
			        if (jQuery("#TxtUserPass").val() == '') {
			            KesionJS.Alert("请输入管理员登录密码！", "jQuery('#TxtUserPass').focus();");
			            return false;
			        }
					if (jQuery("#TxtUserPass").val().length<6) {
			            KesionJS.Alert("密码长度必须大于等于6！", "jQuery('#TxtUserPass').focus();");
			            return false;
			        }
			        if (jQuery("#TxtUserPass").val() != jQuery("#TxtReUserPass").val()) {
			            KesionJS.Alert("两次输入的登录密码不一致，请重输！", "jQuery('#TxtUserPass').focus();");
			            return false;
			        }
			        return true;
			     }
		</script>     
        </div>
	   <%
	   End Sub
	   
	   
	    '删除初始数据
	   sub DelDefaultData(conn)
	     Dim NoDelTable:NoDelTable="KS_Admin,KS_Config,KS_AskGrade,KS_BlogTemplate,KS_Channel,KS_MediaServer,KS_MovieParam,KS_PaymentPlat,KS_Delivery,KS_Deliverytype,KS_PaymentType,KS_Province,KS_User,KS_UserForm,KS_UserGroup,KS_Field,KS_FieldGroup"
		 dim rs,tablename,temptable,tablearr
		 Set rs = Conn.OpenSchema(4)
				tablename=""
				Do Until rs.EOF
					temptable=rs("Table_name")
					if lcase(temptable)<>"ks_notdown" and lcase(left(temptable,3))="ks_"  and foundinarr(lcase(NoDelTable),lcase(temptable),",")=false then
							 Conn.Execute("Delete From " & temptable)
						Tablename = temptable
					end if
				rs.MoveNext
				Loop
				rs.close
		  
	   End Sub
	   
	   
        
	   
	   
	   Sub box_5()
	   dim Conn_Str,DBlx:DBlx=Request("DBlx")
	   if DBlx="1" then  'SQL数据库
	   		Conn_Str="Provider = Sqloledb; User ID = " & Request("TxtDBUser") & "; Password = " & Request("TxtDBPass") & "; Initial Catalog = " & Request("TxtDBName") & "; Data Source = " & Request("TxtDBService") & ";"
			Call OpenConn_install(Conn_Str,1)
	   else
	   		Conn_Str="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../KS_Data/"&Request("TxtDBName_a"))
			Call OpenConn_install(Conn_Str,0)
		    if Request("CkbData")<>"1" then '删除初始数据
			 DelDefaultData(conn)
		    end if
	   end if
	   
	   
	   dim TxtSiteName,TxtSiteUrl,TxtInstallDir,TxtSiteKey,TxtManageDir,TxtManageCode,TxtUserName,TxtUserPass,TxtReUserPass,n,WebSetting	   
	   Dim RSD:Set RSD=Server.CreateObject("ADODB.RECORDSET")
	   RSD.Open "select * from KS_Config",conn,1,3
	   Dim Setting:Setting=Split(RSD("Setting")&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","^%^")
	   
	  
	   dim Loginstr,Loginstr_t,Loginstr_r
	   Loginstr=ReadFromFile("../admin/Login.asp")
	   Loginstr_t=CutFixContent(Loginstr,"'---ShowVerifyCode_s---", "'---ShowVerifyCode_e---",0)
	   Loginstr_r=Loginstr_r&vbcrlf&" Const ShowVerifyCode= "& Request("isCode_a") &"    '后台登录是否启用验证码 true 启用 false不启用"&vbcrlf 
	   Loginstr=Replace(Loginstr,Loginstr_t,Loginstr_r)
	   call WriteTOFile("../admin/Login.asp",Loginstr)
	   
	   call setName("../admin/",Replace(Request("TxtManageDir"),"/",""),1)
	   For n=0 To 240
	   			 select case n
				 case 0
				 	WebSetting=WebSetting &  Request("TxtSiteName") &"^%^" 
				 case 2
				 	WebSetting=WebSetting &  Request("TxtSiteUrl") &"^%^" 
				 case 3
				 	WebSetting=WebSetting &  Request("TxtInstallDir") &"^%^" 
				 case 17
				 	WebSetting=WebSetting &  Request("TxtSiteKey") &"^%^" 
				 case 89
				 	WebSetting=WebSetting &  Request("TxtManageDir") &"^%^" 
				 case else
				 	WebSetting=WebSetting & Setting(n) &"^%^"
				 end select
	   Next
	   RSD("Setting")=WebSetting
	   RSD.Update
	   RSD.Close : Set RSD=Nothing
	   If Trim(Request("TxtUserPass")) <> Trim(Request("TxtReUserPass")) Then
				Response.Write ("<Script>alert('两次输入的登录密码不一致!');history.back(-1);</script>")
				Exit Sub
				Response.End
	   Else
				Conn.Execute("Update KS_User Set [LoginTimes]=0,[UserName]='" & G("TxtUserName") & "',[PassWord]='" &MD5(R(Trim(G("TxtReUserPass"))),16) &"' Where UserName='admin'")  
				Conn.Execute("Update KS_Admin Set [LoginTimes]=0,[UserName]='" & G("TxtUserName") & "',[PrUserName]='" & G("TxtUserName") & "',[PassWord]='" &MD5(R(Trim(G("TxtReUserPass"))),16) &"' Where UserName='admin'")  
				Conn.Execute("Delete From KS_User  Where UserName<>'" & G("TxtUserName") & "'")  
				Conn.Execute("Delete From KS_Admin  Where UserName<>'" & G("TxtUserName") & "'")  
	   End If
	   dim str_co,str_t,str_reco
	   str_co=ReadFromFile("../conn.asp")

	   str_t=CutFixContent(str_co,"'========认证密码开始========", "'========认证密码结束========",0)
	   str_reco= str_reco &vbcrlf&"Const EnableSiteManageCode = "& Request("isCode") &"        '是否启用后台管理认证密码 是： True  否： False " &vbcrlf 
	   str_reco= str_reco &"Const SiteManageCode = """& Request("TxtManageCode") &"""      '后台管理认证密码，请修改，这样即使有人知道了您的后台用户名和密码也不能登录后台"&vbcrlf
	   str_co=Replace(str_co,str_t,str_reco)
	   
	   str_reco=""
	   str_t=CutFixContent(str_co,"'========数据库类型开始========", "'========数据库类型结束========",0)
	   str_reco= str_reco &vbcrlf&"Const DataBaseType="& Request("DBlx") &"                 '系统数据库类型，""1""为MS SQL2000数据库，""0""为MS ACCESS 2000数据库" &vbcrlf 
	   str_co=Replace(str_co,str_t,str_reco)
	   
	   str_reco=""
	   str_t=CutFixContent(str_co,"'========数据库设置开始========", "'========数据库设置结束========",0)
	   str_reco= str_reco &vbcrlf&" If DataBaseType=0 then"&vbcrlf 
	   str_reco= str_reco &	"'如果是ACCESS数据库，请认真修改好下面的数据库的文件名"&vbcrlf 
	   str_reco= str_reco &	"	DBPath       = """& "/KS_Data/"&Request("TxtDBName_a") &"""     'ACCESS数据库的文件名，请使用相对于网站根目录的的绝对路径"&vbcrlf 
	   str_reco= str_reco &"Else"&vbcrlf 
	   str_reco= str_reco &"		 '如果是SQL数据库，请认真修改好以下数据库选项"&vbcrlf 
	   str_reco= str_reco &"	 DataServer   = """& Request("TxtDBService") &"""                                  '数据库服务器IP"&vbcrlf 
	   str_reco= str_reco &"	 DataUser     = """& Request("TxtDBUser") &"""                                       '访问数据库用户名"&vbcrlf 
	   str_reco= str_reco &"	 DataBaseName = """& Request("TxtDBName") & """                                '数据库名称"&vbcrlf 
	   str_reco= str_reco &"	 DataBasePsw  = """& Request("TxtDBPass") &"""                                   '访问数据库密码"&vbcrlf 
	   str_reco= str_reco &"End if"&vbcrlf 
	   str_co=Replace(str_co,str_t,str_reco)
	   
	   
	   call WriteTOFile("../Conn.asp",str_co)
	   call WriteTOFile("install.lock","lock")
	   
	   %>
       
       
	   <div id="Step5">
	
		<div class="datiao">
			<ul>
				<li>阅读许可协议</li>
				<li>检查安装环境</li>
				<li>创建数据库</li>
				<li>网站参数配置</li>
				<li class="curr">完成安装</li>
			</ul>
		 </div>
		<div class="clear"></div>
		
		<div class="hjlist" style="padding:20px"><br><br>
		  <h4 style="color:Green;line-height:35px;"><img src="images/ok.png" align="absmiddle">恭喜，KesionCMS X2.0产品安装成功！
		  <br>为了您的网站安全，建议及时删除<font color="red">“install”</font>安装目录。</h4>
		  <div style="padding:10px"><br><br>
		  <a href="<%=Request("TxtInstallDir")%><%=Request("TxtManageDir")%>Login.asp">进入后台</a> | <a href="../">进入网站前台</a> | <a href="http://www.kesion.com/" target="_blank">浏览官方网站</a> | <a href="http://bbs.kesion.com/" target="_blank">浏览技术论坛</a>
		  </div>
		  <br><br><br><br>
		  </div>
                
        </div>
	   <%
	   End Sub
		
		Function CutFixContent(ByVal str, ByVal start, ByVal last, ByVal n)
			Dim strTemp
			On Error Resume Next
			If InStr(str, start) > 0 Then
				Select Case n
				Case 0  '左右都截取（都取前面）（去处关键字）
					strTemp = Right(str, Len(str) - InStr(str, start) - Len(start) + 1)
					strTemp = Left(strTemp, InStr(strTemp, last) - 1)
				Case Else  '左右都截取（都取前面）（保留关键字）
					strTemp = Right(str, Len(str) - InStr(str, start) + 1)
					strTemp = Left(strTemp, InStr(strTemp, last) + Len(last) - 1)
				End Select
			Else
				strTemp = ""
			End If
			CutFixContent = strTemp
		End Function
		Public Function ISFromFile(FileName)
		   On Error Resume Next
			dim str,stm
			set stm=server.CreateObject("adodb.stream")
			stm.Type=2 '以本模式读取
			stm.mode=3 
			stm.charset="utf-8"
			stm.open
			stm.loadfromfile server.MapPath(FileName)
			str=stm.readtext
			stm.Close
			set stm=nothing
			if err.number<>0 then 
				ISFromFile="Error" :Exit Function
			else
				ISFromFile="ok" :Exit Function
			end if
		End Function
		
		Public Function CheckDir(FolderPath)
	        Dim fso:Set fso = server.createobject("Scripting.FileSystemObject")
			CheckDir=fso.FolderExists(Server.MapPath(FolderPath))
			Set fso = Nothing
	   End Function
	
		sub Copy(filename,FolderName,cs)
			dim Fso,MyFile
			set Fso=server.createobject("Scripting.FileSystemObject")
			filename=server.MapPath(filename)
			FolderName=server.MapPath(FolderName)
			if cs=0 then
				Set MyFile = fso.GetFile(filename)
			else
				Set MyFile = fso.GetFolder(filename)
			end if
			MyFile.Copy FolderName
		end sub
		sub setName(filename,name,cs)
		   if lcase(name)="admin" then exit sub
			dim Fso,MyFile
			set Fso=server.createobject("Scripting.FileSystemObject")
			filename=server.MapPath(filename)
			if cs=0 then
			Set MyFile = Fso.GetFile(filename)
			else
			Set MyFile = Fso.GetFolder(filename)
			end if
			MyFile.Name=name
		end sub
		sub FileDel(filename,cs)
			dim Fso,MyFile
			set Fso=server.createobject("Scripting.FileSystemObject")
			filename=server.MapPath(filename)
			if cs=0 then
			Set MyFile = fso.GetFile(filename)
			else
			Set MyFile = fso.GetFolder(filename)
			end if
			MyFile.Delete
		end sub
		function fileExists(filename)
			dim Fso:set Fso=server.createobject("Scripting.FileSystemObject")
			if fso.fileExists(server.mappath(filename)) Then
			fileExists=true
			else
			fileExists=false
			end if
		end function
		
	'**************************************************
	'函数名：WriteTOFile
	'作  用：写内容到指定的html文件
	'参  数：Filename  ----目标文件件 如 mb\index.htm
	'        Content   ------要写入目标文件的内容
	'返回值：成功返回true ,失败返回false
	'**************************************************
	Public Function WriteTOFile(FileName, Content)
	   On Error Resume Next
		dim stm:set stm=server.CreateObject("adodb.stream")
		stm.Type=2 '以文本模式读取
		stm.mode=3
		stm.charset="utf-8"
		stm.open
		stm.WriteText content
		stm.SaveToFile server.MapPath(FileName),2 
		stm.flush
		stm.Close
		set stm=nothing
	   If Err.Number <> 0 Then
		 WriteTOFile = False
	   Else
		 WriteTOFile = True
	   End If
	End Function
	
	'**************************************************
	'函数名：ReadFromFile
	'作  用：写内容到指定的html文件
	'参  数：Filename  ----目标文件件 如 mb\index.htm
	'返回值：成功返回文件内容 ,失败返回""
	'**************************************************
	Public Function ReadFromFile(FileName)
	    On Error Resume Next
		dim str,stm
		set stm=server.CreateObject("adodb.stream")
		stm.Type=2 '以本模式读取
		stm.mode=3 
		stm.charset="utf-8"
		stm.open
		stm.loadfromfile server.MapPath(FileName)
		str=stm.readtext
		stm.Close
		set stm=nothing
		if err.number<>0 then Die ("Error:<br/>[" & Server.MapPath(FileName) & "] File does not exist!"):Exit Function
		ReadFromFile=str
	End Function
				
		function isFsowrite(filename)
			On Error Resume Next
			call WriteTOFile(filename,"a")
			Call FileDel(filename,0)
			if err=0 then
				isFsowrite= "<img src=""images/v.png"" align=""absmiddle"">可写"
			else
				isFsowrite= "<img src=""images/no.gif"" align=""absmiddle"">不可写"
				IsCheckEnvironmentPass=false
			end if	  
		end function
		
		
End Class
%>


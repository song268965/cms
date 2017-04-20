<%
Sub CreateJs(id)
	dim param:param=" where 1=1"
	If ID<>0 Then param=param & " and place=" & id
	dim rs:set rs=server.createobject("adodb.recordset")
	rs.open "select * from KS_ADPlace" & param,conn,1,1
	if not rs.eof then
	    do while not rs.eof
				 dim rst:set rst=server.createobject("adodb.recordset")
				 dim str,i,placeId,SaveFilePath,placelei,placewid,placehei
				 i=0 : placeId=rs("place") :placelei=rs("placelei") : str=""
				placehei=rs("placehei")
				placewid=rs("placewid")
				
				GaoAndKuan=""
				
				if Not KS.IsNUL(placehei) then GaoAndKuan=" height="&placehei&" "
				if Not KS.IsNul(placewid) then GaoAndKuan=GaoAndKuan&" width="&placewid&" "	
				
				dim sqlstr:sqlstr="select * from KS_Advertise where act=1 and place="& placeId & " order by AdOrderID,id"
				 rst.open sqlstr,conn,1,1
				 select case placelei
				   case 1
					 str="document.write(""<span id='s" &placeId & "'></span>"");" & vbcrlf
					 str=str & "var GetRandomn = 1;" & vbcrlf
					 str=str & "function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}" & vbcrlf
					 str=str &" var a" & placeId & "=new Array();" & vbcrlf
					 str=str & "var t"&placeId & "=new Array();" &vbcrlf
					 str=str & "var ts" & placeId &"=new Array();" & vbcrlf
					 str=str & "var allowclick" & placeId &"=new Array();" & vbcrlf
					 str=str & "var id" & placeId &"=new Array();" & vbcrlf
					 do while not rst.eof
					   if rst("xslei")="swf" then
					    str=str & "a" & placeId & "[" & i & "]=""" & DggtXs(rst) & """;" & vbcrlf
					   else
					    str=str & "a" & placeId & "[" & i & "]=""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span>"";" & vbcrlf
					   end if
					   str=str & "t" & placeId&"["&i&"]=" & rst("class") &";" &vbcrlf
					   str=str & "ts" & placeId&"["&i&"]=""" & year(rst("lasttime"))&"-" & month(rst("lasttime"))&"-" & day(rst("lasttime")) & """;" & vbcrlf
					   str=str & "allowclick" & placeId & "["&i&"]=" & KS.ChkClng(rst("allowclick")) & ";" &vbcrlf
					   str=str & "id" & placeId & "["&i&"]=" & rst("id") & ";" &vbcrlf
					   i=i+1
					  rst.movenext
					 loop
					 str=str & "var temp" & placeId & "=new Array();" &vbcrlf
					 str=str & "var k=0;" & vbcrlf
					 str=str & "for(var i=0;i<a" & placeId &".length;i++){" &vbcrlf
					 str=str & "if (t" & placeId &"[i]==1){" & vbcrlf
					 str=str & "if (checkDate"&placeId&"(ts" & placeId&"[i])){" &vbcrlf
					 str=str &"	temp"& placeId&"[k++]=a" &placeId&"[i];" & vbcrlf & "}"&vbcrlf
					 str=str &"	}else{"&vbcrlf
					 str=str &" temp" & placeID&"[k++]=a" & placeID&"[i];" & vbcrlf &"}"&vbcrlf
					 str=str & "}" & vbcrlf
					 
					 str=str & "if (temp"&placeId & ".length>0){"&vbcrlf
					 str=str & "GetRandom(temp" & placeId & ".length);" & vbcrlf
					 str=str & "var index" & placeId & "=GetRandomn-1;" & vbcrlf
					 str=str & "if (allowclick" &placeId &"[index" & placeId & "]>0){ " & vbcrlf
					 str=str & "jQuery.getScript('" & KS.GetDomain &"plus/ads/showA.asp?action=loadjs&times='+allowclick" &placeId &"[index" & placeId & "]+'&id='+id" &placeId &"[index" & placeId & "],function(){  " &vbcrlf
					 str=str & "$('#s" & placeId &"').html(a" &placeId &"[index" & placeId & "]);" & vbcrlf
					 str=str & "if (data.isEnd=='1'){" &vbcrlf
					 str=str & GetEndTips(placeId) &vbcrlf
					 str=str & " } });" & vbcrlf
					 str=str & "}else{" & vbcrlf
					 str=str & "$('#s" & placeId &"').html(a" &placeId &"[index" & placeId & "]);" & vbcrlf
					 str=str & "}" & vbcrlf
					 str=str & "}"&vbcrlf
					 str=str & getClicks(placeId)
					 
				  case 2
				    str="document.write(""<span id='s" &placeId & "'></span>"");" & vbcrlf
				   do while not rst.eof
				     str=str &"if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
					  
					  if KS.ChkClng(rst("allowclick"))>0 Then  '限制每天点击次数
					  
					     str=str & "jQuery.getScript('" & KS.GetDomain &"plus/ads/showA.asp?action=loadjs&times=" & KS.ChkClng(rst("allowclick")) &"&id=" &rst("id") &"',function(){  " &vbcrlf
						 
						 if rst("xslei")="swf" then
						   str=str & "$('#s" & placeId &"').html(""" & DggtXs(rst) & "<br/>"");" & vbcrlf
						 else
						   str=str & "$('#s" & placeId &"').html(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span><br/>"");" & vbcrlf
						 end if
						 
						 str=str & "if (data.isEnd=='1'){" &vbcrlf
						 str=str & GetEndTips(placeId) &vbcrlf
						 str=str & " } });" & vbcrlf
					  
					  
					  Else
						 if rst("xslei")="swf" then
						  str=str &"document.writeln(""" & DggtXs(rst) & "<br/>"");" & vbcrlf
						 else
						  str=str &"document.writeln(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span><br/>"");" & vbcrlf
						 end if
	                 End If				 
					 
					 str=str & "}" &vbcrlf
					rst.movenext
				   loop
					 str=str & getClicks(placeId)
				  case 3
				   str="document.write(""<span id='s" &placeId & "'></span>"");" & vbcrlf
				   do while not rst.eof
				     str=str &"if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
					 
					 
					 if KS.ChkClng(rst("allowclick"))>0 Then  '限制每天点击次数
					     str=str & "jQuery.getScript('" & KS.GetDomain &"plus/ads/showA.asp?action=loadjs&times=" & KS.ChkClng(rst("allowclick")) &"&id=" &rst("id") &"',function(){  " &vbcrlf
						 
						 if rst("xslei")="swf" then
						   str=str & "$('#s" & placeId &"').html(""" & DggtXs(rst) & "&nbsp;"");" & vbcrlf
						 else
						   str=str & "$('#s" & placeId &"').html(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span>&nbsp;"");" & vbcrlf
						 end if
						 
						 str=str & "if (data.isEnd=='1'){" &vbcrlf
						 str=str & GetEndTips(placeId) &vbcrlf
						 str=str & " } });" & vbcrlf
					 
					 
					 Else
						 if rst("xslei")="swf" then
						  str=str &"document.write(""" & DggtXs(rst) & "&nbsp;"");" & vbcrlf
						 else
						  str=str &"document.write(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span>&nbsp;"");" & vbcrlf
						 end if 
					End If	 
					 
					 
					 str=str &"}" &vbcrlf
					rst.movenext
				   loop
					 str=str & getClicks(placeId)
				  case 4
				   str="document.write('<marquee  direction=""up"""&GaoAndKuan&">');" & vbcrlf
				   str=str & "document.write(""<span id='s" &placeId & "'></span>"");" & vbcrlf
				   do while not rst.eof
				     str=str &"if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
					  IF KS.ChkClng(rst("allowclick"))>0 Then  '限制每天点击次数
					  
					     str=str & "jQuery.getScript('" & KS.GetDomain &"plus/ads/showA.asp?action=loadjs&times=" & KS.ChkClng(rst("allowclick")) &"&id=" &rst("id") &"',function(){  " &vbcrlf
						 
						 if rst("xslei")="swf" then
						   str=str & "$('#s" & placeId &"').html(""" & DggtXs(rst) & "<br/><br/>"");" & vbcrlf
						 else
						   str=str & "$('#s" & placeId &"').html(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span><br/><br/>"");" & vbcrlf
						 end if
						 
						 str=str & "if (data.isEnd=='1'){" &vbcrlf
						 str=str & GetEndTips(placeId) &vbcrlf
						 str=str & " } });" & vbcrlf
					 
					  
					  Else
						 if rst("xslei")="swf" then
						 str=str &"document.write(""" & DggtXs(rst) & "<br/><br/>"");" & vbcrlf
						 else
						 str=str &"document.write(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span><br/><br/>"");" & vbcrlf
						 end if
					 End If
					 str=str &"}" &vbcrlf
					rst.movenext
				   loop
				   str=str &"document.write(""</marquee>"");" & vbcrlf
				   str=str & getClicks(placeId)
				  case 5
				   str="document.write('<marquee"&GaoAndKuan&">');" & vbcrlf
				   str=str & "document.write(""<span id='s" &placeId & "'></span>"");" & vbcrlf
				   do while not rst.eof
				     str=str &"if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
					 IF KS.ChkClng(rst("allowclick"))>0 Then  '限制每天点击次数
					     str=str & "jQuery.getScript('" & KS.GetDomain &"plus/ads/showA.asp?action=loadjs&times=" & KS.ChkClng(rst("allowclick")) &"&id=" &rst("id") &"',function(){  " &vbcrlf
						 
						 if rst("xslei")="swf" then
						   str=str & "$('#s" & placeId &"').html(""" & DggtXs(rst) & "&nbsp;"");" & vbcrlf
						 else
						   str=str & "$('#s" & placeId &"').html(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span>&nbsp;"");" & vbcrlf
						 end if
						 
						 str=str & "if (data.isEnd=='1'){" &vbcrlf
						 str=str & GetEndTips(placeId) &vbcrlf
	
						 str=str & " } });" & vbcrlf
					 Else
						 if rst("xslei")="swf" then
						 str=str &"document.write(""" & DggtXs(rst) & "&nbsp;"");" & vbcrlf
						 else
						 str=str &"document.write(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span>&nbsp;"");" & vbcrlf
						 end if
					End If
					 str=str &"}" &vbcrlf
					rst.movenext
				   loop
				   str=str &"document.write(""</marquee>"");" & vbcrlf
				   str=str & getClicks(placeId)
				  case 6
				   do while not rst.eof
					 call gaokuan(rs,rst)
					 str=str &"if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
					str=str & "window.open('"&DomainStr&"plus/ads/ShowA.asp?Action=AdOpen&i="&rst("id")&"','" & KS.Setting(0) & "广告服务"&rst("id")&"','"&GaoAndKuan&"');" &vbcrlf
					str=str &"}" &vbcrlf
		
					rst.movenext
				   loop
				   str=str & getClicks(placeId)
				  case 7
					 str="var GetRandomn = 1;" & vbcrlf
					 str=str & "function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}" & vbcrlf
					 str=str &" var a" & placeId & "=new Array();" & vbcrlf
					 str=str &" var gk" & placeId & "=new Array();" & vbcrlf
					 str=str & "var t"&placeId & "=new Array();" &vbcrlf
					 str=str & "var ts" & placeId &"=new Array();" & vbcrlf
					 do while not rst.eof
					   str=str & "t" & placeId&"["&i&"]=" & rst("class") &";" &vbcrlf
					   str=str & "ts" & placeId&"["&i&"]=""" & formatdatetime(rst("lasttime"),2) & """;" & vbcrlf
					   str=str & "a" & placeId & "[" & i & "]="""&DomainStr&"plus/ads/ShowA.asp?Action=AdOpen&i="&rst("id")&""";" & vbcrlf
					   call gaokuan(rs,rst)
					   str=str & "gk" & placeId & "[" & i & "]="""&GaoAndKuan&""";" & vbcrlf
					   i=i+1
					  rst.movenext
					 loop
					 str=str & "var temp" & placeId & "=new Array();" &vbcrlf
					 str=str & "var k=0;" & vbcrlf
					 str=str & "for(var i=0;i<a" & placeId &".length;i++){" &vbcrlf
					 str=str & "if (t" & placeId &"[i]==1){" & vbcrlf
					 str=str & "if (checkDate"&placeId&"(ts" & placeId&"[i])){" &vbcrlf
					 str=str &"	temp"& placeId&"[k++]=a" &placeId&"[i];" & vbcrlf & "}"&vbcrlf
					 str=str &"	}else{"&vbcrlf
					 str=str &" temp" & placeID&"[k++]=a" & placeID&"[i];" & vbcrlf &"}"&vbcrlf
					 str=str & "}" & vbcrlf
					 str=str & "if (temp"&placeId & ".length>0){"&vbcrlf
					 str=str & "GetRandom(temp" & placeId & ".length);" & vbcrlf
					 str=str & "window.open(temp" &placeId &"[GetRandomn-1],'"&KS.Setting(0)&"广告服务',gk"&PlaceId&"[GetRandomn-1]);" & vbcrlf
		             str=str & "}"&vbcrlf
					 str=str & getClicks(placeId)
				 case 8,9 '对联广告
				     dim nn:nn=0
					dim left1:left1=""
					dim left2:left2=""
				   do while not rst.eof
					  
					     if nn=0 then
				            left1="if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
							 if rst("xslei")="swf" then
							  left1=left1 &"document.writeln(""" & DggtXs(rst) & "<br/>"");" & vbcrlf
							 else
							  left1=left1 & "document.writeln(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span><br/>"");" & vbcrlf
							 end if
							 left1=left1 & "}" &vbcrlf

					     else
						      left2="if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
							 if rst("xslei")="swf" then
							  left2=left2 & "document.writeln(""" & DggtXs(rst) & "<br/>"");" & vbcrlf
							 else
							  left2=left2 & "document.writeln(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span><br/>"");" & vbcrlf
							 end if
							 left2=left2 & "}" &vbcrlf
						 end if
					 nn=nn+1
					 if nn>=2 then exit do
					rst.movenext
				   loop
				 
				    str=str & "lastScrollY = 0;" &vbcrlf
					str=str & "function heartBeat(){" &vbcrlf
					str=str & "var diffY;" &vbcrlf
					str=str & "if (document.documentElement && document.documentElement.scrollTop)" &vbcrlf
					str=str & "diffY = document.documentElement.scrollTop;" &vbcrlf
					str=str & "else if (document.body)" &vbcrlf
					str=str & "diffY = document.body.scrollTop" &vbcrlf
					str=str & "else" &vbcrlf
					str=str & "{/*Netscape stuff*/}" &vbcrlf
					str=str & "percent=.1*(diffY-lastScrollY);" &vbcrlf
					str=str & "if(percent>0)percent=Math.ceil(percent);" &vbcrlf
					str=str & "else percent=Math.floor(percent);" &vbcrlf
					str=str & "document.getElementById(""leftDiv"").style.top = parseInt(document.getElementById(""leftDiv"").style.top)+percent+""px"";" &vbcrlf
					str=str & "document.getElementById(""rightDiv"").style.top = parseInt(document.getElementById(""rightDiv"").style.top)+percent+""px"";" &vbcrlf
					str=str & "lastScrollY=lastScrollY+percent;" &vbcrlf
					str=str & "}" &vbcrlf
					
					if placelei=8 then
					str=str & "//下面这段删除后，对联将不跟随屏幕而移动。" &vbcrlf
					str=str & "window.setInterval(""heartBeat()"",1);" &vbcrlf
					end if
					
					str=str & "//-->" &vbcrlf
					str=str & "//关闭按钮" &vbcrlf
					str=str & "function close_left1(){left1.style.visibility='hidden';}" &vbcrlf
					str=str & "function close_right1(){right1.style.visibility='hidden';}" &vbcrlf
					
					str=str & "//显示样式" &vbcrlf
					str=str & "document.writeln(""<style type=\""text\/css\"">"");" &vbcrlf
					str=str & "document.writeln(""#leftDiv,#rightDiv{width:" & placewid &"px;height:" & placehei &"px;background-color:#fff;position:absolute;}"");" &vbcrlf
					str=str & "document.writeln("".itemFloat{width:" & placewid &"px;height:auto;line-height:5px}"");" &vbcrlf
					str=str & "document.writeln("".itemFloat img{width:" & placewid &"px;height:" & placehei &"px;}"");" &vbcrlf
					str=str & "document.writeln(""<\/style>"");" &vbcrlf
					str=str & "//以下为主要内容" &vbcrlf
					str=str & "document.writeln(""<div id=\""leftDiv\"" style=\""top:100px;left:5px\"">"");" &vbcrlf
					str=str & "//------左侧各块开始" &vbcrlf
					str=str & "//---L1" &vbcrlf
					str=str & "document.writeln(""<div id=\""left1\"" class=\""itemFloat\"">"");" &vbcrlf
					
					str=str & left1
					
					
					str=str & "document.writeln(""<br><a href=\""javascript:close_left1();\"" title=\""关闭上面的广告\"">×<\/a><br><br><br><br>"");" &vbcrlf
					str=str & "document.writeln(""<\/div>"");" &vbcrlf
					
					str=str & "//------左侧各块结束" &vbcrlf
					str=str & "document.writeln(""<\/div>"");" &vbcrlf
					str=str & "document.writeln(""<div id=\""rightDiv\"" style=\""top:100px;right:5px\"">"");" &vbcrlf
					str=str & "//------右侧各块结束" &vbcrlf
					str=str & "//---R1" &vbcrlf
					str=str & "document.writeln(""<div id=\""right1\"" class=\""itemFloat\"">"");" &vbcrlf
					
					str=str & left2
					str=str & "document.writeln(""<br><a href=\""javascript:close_right1();\"" title=\""关闭上面的广告\"">×<\/a><br><br><br><br>"");" &vbcrlf
					str=str & "document.writeln(""<\/div>"");" &vbcrlf
					
					str=str & "//------右侧各块结束" &vbcrlf
					str=str & "document.writeln(""<\/div>"");" &vbcrlf
					 str=str & getClicks(placeId)

				 end select	
				   rst.close : set rst=nothing
				 SaveFilePath = KS.Setting(3) & KS.Setting(93) 
				 Call KS.CreateListFolder(SaveFilePath)
				 if KS.ChkClng(rs("show_flag"))=1 then
				 Call KS.WriteTOFile(SaveFilePath& placeId & ".js", str)
				 else
				 Call KS.WriteTOFile(SaveFilePath& placeId & ".js","document.write('');")
				 end if
		  RS.MoveNext
	   Loop
	end if
	rs.close
	set rs=nothing
    
  End Sub
  
  function GetEndTips(placeId)
   	 Dim str:str="$('#s" & placeId &"').find(""a"").attr(""href"",""javascript:;"").click(function(){ alert(""对不起，该广告今天已达到点击上限！""); return false; });" & vbcrlf
     str=str & "$('#s" & placeId &"').find(""span"").click(function(){return false; });" & vbcrlf
    GetEndTips=str
  end Function
  
  function getClicks(placeId)
   Dim str
   str="function addHits" & placeId&"(c,id){if(c==1){try{jQuery.getScript('" & domainStr &"plus/ads/showa.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}" & vbcrlf
   str=str & "function checkDate" & placeId&"(date_arr){" &vbcrlf
   str=str &" var date=new Date();" &vbcrlf
   str=str &" date_arr=date_arr.replace(/\//g,""-"").split(""-"");" &vbcrlf
   str=str & "var year=parseInt(date_arr[0]);" & vbcrlf
   str=str & "var month=parseInt(date_arr[1])-1;" & vbcrlf
   str=str & "var day=0;" & vbcrlf
   str=str & "if (date_arr[2].indexOf("" "")!=-1)" & vbcrlf
   str=str & "day=parseInt(date_arr[2].split("" "")[0]);" & vbcrlf
   str=str & "else" & vbcrlf
   str=str & "day=parseInt(date_arr[2]);" &vbcrlf
   str=str & "var date1=new Date(year,month,day);" & vbcrlf
   str=str & "if(date.valueOf()>date1.valueOf())" & vbcrlf
   str=str &" return false;" &vbcrlf
   str=str &"else" &vbcrlf
   str=str &" return true" & vbcrlf
   str=str &"}" &vbcrlf
   getClicks=str
  end function
  
  Function DggtXs(rst)
    dim str,ttarg,GaoAndKuan,GKCss
	if rst("window")=0 then
		ttarg = "_blank"
	else 
		ttarg = "_top" 
	end if
    if isnumeric(rst("hei")) then
		GaoAndKuan=" height="&rst("hei")&" "
		GKCss="height:" &rst("hei")&"px;"
	else
		
		if right(rst("hei"),1)="%" then
		if isnumeric(Left(rst("hei"),len(rst("hei"))-1))=true then
		 GaoAndKuan=" height="&rst("hei")&" "
		 GKCss="height:" &rst("hei")&";"
		end if
		end if
		
		end if
		
		
		if isnumeric(rst("wid")) then
		GaoAndKuan=GaoAndKuan&" width="&rst("wid")&" "
		GKCss=GKCss&"width:" &rst("wid")&"px;"
		else
		if right(rst("wid"),1)="%" then
		if isnumeric(Left(rst("wid"),len(rst("wid"))-1))=true then 
		GaoAndKuan=GaoAndKuan&" width="&rst("wid")&" "
		GKCss=GKCss&"width:" &rst("wid")&";"
		end if
		end if
	end if	
	 dim gif_url:gif_url=rst("gif_url")
	 if left(lcase(gif_url),4)<>"http" then gif_url=KS.Setting(2) & gif_url
     Select Case rst("xslei")
		   Case "txt"
		    str="<a title=""" & rst("sitename") & """  href=""" & rst("url") & """ target=""" & ttarg & """>" & rst("sitename") & "</a>"
		   Case "gif"
		    str="<a href=""" &  rst("url") & """ target=""" & ttarg & """><img  alt=""" & rst("sitename") & """  border=""0"" " & GaoAndKuan&" src=""" & gif_url & """></a>"
		   Case "swf"
		   str="<a href=""" & rst("url") & """ onclick=""addHits" & rst("place")&"(" & rst("clicks") &"," & rst("id") & ")"" target=""" & ttarg & """ hidefocus><button disabled style=""margin:0px;padding:0px;cursor:pointer;border:none;" &GKCss &"""><object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" "&GaoAndKuan &">"
		   str=str & "<param name=""movie"" value=""" & gif_url &""" />"
		   str=str & "<param name=""quality"" value=""high"" />"
		   str=str & "<param name=""wmode"" value=""transparent"" />"
		   str=str & "<embed src=""" & gif_url & """ quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" " &GaoAndKuan &"></embed>"
		   str=str & "</object></button></a>"
		   Case "dai"
		    str="<iframe marginwidth=""0"" marginheight=""0""  frameborder=""0"" bordercolor=""000000"" scrolling=""no""  name=""广告"" src=""" & DomainStr & "plus/ads/ShowA.asp?Action=Daima&id=" & rst("id") & """  " & GaoAndKuan &"></iframe>"
		  Case else
		    str="<a href=""" & rst("url") & """ target=""" & ttarg & """><img alt=""" & rst("sitename") & """  border=""0"" " & GaoAndKuan &" src=""" & gif_url & """ /></a>"
	End Select
	str=Replace(Replace(Replace(Replace(str, Chr(13)& Chr(10), ""),"'","\'"),"""","\"""),vbcrlf,"") 
	DggtXs=str	
  End Function
  
  
Sub gaokuan(rs,adsrs) 
		if not KS.IsNul(adsrs("hei")) and adsrs("hei")<>"0" then
			if isnumeric(adsrs("hei")) then
			  GaoAndKuan="height="&adsrs("hei")
			else
				 if right(adsrs("hei"),1)="%" then
				   if isnumeric(Left(len(adsrs("hei"))-1))=true then
					 GaoAndKuan="height="&adsrs("hei")
				   end if
				 end if
			end if
		else
		  GaoAndKuan="height="&rs("placehei")
		end if
		
	  If Not KS.IsNul(adsrs("wid")) and adsrs("wid")<>"0" Then
		if isnumeric(adsrs("wid")) then
		   GaoAndKuan=GaoAndKuan&",width="&adsrs("wid")
		else
			if right(adsrs("wid"),1)="%" then
				if isnumeric(Left(len(adsrs("wid"))-1))=true then 
				 GaoAndKuan=GaoAndKuan&",width="&adsrs("wid")
				end if
			end if
		end if
	  Else
	    GaoAndKuan=GaoAndKuan&",width="&rs("placewid")
	  End If
	End Sub
%>
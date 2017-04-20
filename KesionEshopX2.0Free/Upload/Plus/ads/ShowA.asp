<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls = New Advertise
KSCls.Kesion()
Set KSCls = Nothing

Class Advertise
        Private KS
		Private getplace,getshow,adsrs,adssql,adsrsp,adssqlp,adsrss,adssqls,getip,getggwlxsz,getggwhei,getggwwid
        Private ttarg,DomainStr,GaoAndKuan,advertvirtualvalue
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
		%>
		<!--#include file="CreateAdsFun.asp"-->
		<%
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		  Select Case KS.S("Action")
		   Case "loadjs"
		     Call loadjs()
		   Case "Daima"
		     Call AdvertiseDaima()
		   Case "AdOpen"
		     Call AdvertiseAdOpen()
		   Case "HitsGuangGao" 
		     Call HitsGuangGao()
		  End Select
		End Sub
  Sub loadjs()
     Dim placeId:placeId=KS.ChkClng(Request("placeId"))
	 Dim id:id=KS.ChkClng(Request("id"))
	 Dim allowclick:allowclick=KS.ChkClng(Request("times"))
	 if id=0 then ks.die ""
	 If allowclick>0 Then
	       Dim CurrClickNum:CurrClickNum=KS.ChkClng(Conn.Execute("Select count(1) From KS_Adiplist Where adid=" & id & " and DateDiff(" & DataPart_D & ",time," & SQLNowString &")<1")(0))
		   If CurrClickNum >= allowclick Then
		    
		     KS.Die "var data={'isEnd':'1'}"
		   End If
	 End If
	 KS.Die "var data={'isEnd':'0'}"
	
  End Sub
		
 '代码
  Sub AdvertiseDaima()
         response.write "<body>"
  	    if KS.S("id")<>"" and isnumeric(KS.S("ID")) then
			dim adssql
			dim adsrs:set adsrs=server.createobject("adodb.recordset")
			adssql="Select top 1 intro from KS_Advertise where id="&KS.ChkClng(KS.S("id"))&" order by time"
			adsrs.open adssql,conn,1,1       
			if not adsrs.eof then
			response.write adsrs(0)
			end if
			adsrs.close:set adsrs=nothing
			conn.close:set conn=nothing
		else
			response.write "<center><br><br>无效广告。</center>"
		end if
		response.write "</body>"
  End Sub

 Sub AdvertiseAdOpen()
 %>
     <html>
	 <head>
	 <script type="text/javascript" src="../../ks_inc/jquery.js"></script>
	 <script type="text/javascript">
	 function addHits(c,id){if(c==1){try{jQuery.getScript('showa.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
	 </script>
	 <style type="text/css">
	 body{font-size:12px}
	 </style>
	 </head>
	 <body topmargin="0" leftmargin="0">
	<%
	Dim DomainStr:DomainStr=KS.GetDomain
	Dim ttarg:ttarg="_top"
	Dim GaoAndKuan:GaoAndKuan=""
	Dim Adsrs:Set adsrs=server.createobject("adodb.recordset")
	Dim adssql:adssql="Select top 1 id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei,clicks,url,allowclick from KS_Advertise where id="&KS.Chkclng(KS.S("i"))
	adsrs.open adssql,Conn,3,3
	adsrs("show")=adsrs("show")+1
	adsrs("time")=now()
	adsrs.Update
	if adsrs("window")=0 then
	ttarg = "_blank"
	end if
	dim placeId:placeId=adsrs("place")
	Dim allowclick:allowclick=KS.ChkClng(adsrs("allowclick"))
	dim isend:isend=false
	If allowclick>0 Then '判断每天限制次数
		   Dim CurrClickNum:CurrClickNum=KS.ChkClng(Conn.Execute("Select count(1) From KS_Adiplist Where adid=" & KS.Chkclng(KS.S("i")) & " and DateDiff(" & DataPart_D & ",time," & SQLNowString &")<1")(0))
		   If CurrClickNum >= allowclick Then
		     isend=true
		   End If
		End If
		
	if isnumeric(adsrs("hei")) then
	GaoAndKuan=" height="&adsrs("hei")&" "
	else
	
	if right(adsrs("hei"),1)="%" then
		if isnumeric(Left(len(adsrs("hei"))-1))=true then
		 GaoAndKuan=" height="&adsrs("hei")&" "
		end if
	end if
	
	end if
	
	if isnumeric(adsrs("wid")) then
	GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
	else
	if right(adsrs("wid"),1)="%" then
	if isnumeric(Left(len(adsrs("wid"))-1))=true then 
	GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
	end if
	end if
	end if
	
	 Select Case adsrs("xslei")
				Case "txt"%>
				  <%If isend Then%>
				<span onclick="alert('对不起，该广告今天已达到点击上限！');"><a title="<%=adsrs("sitename")%>" href="javascript:;"><%=adsrs("sitename")%></a></span>
				  <%else%>
				<span onClick="addHits(<%=adsrs("clicks")%>,<%=adsrs("id")%>)"><a title="<%=adsrs("sitename")%>" href="<%=adsrs("url")%>" target="<%=ttarg%>"><%=adsrs("sitename")%></a></span>
				  <%end if%>
	<%          Case "gif"%>
	              <%If isend Then%>
				    <span onclick="alert('对不起，该广告今天已达到点击上限！');">
	                <a title="<%=adsrs("intro")%>" href="javascript:;"><img border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>" /></a> 
				    </span>
				  <%else%>
	                <span onClick="addHits(<%=adsrs("clicks")%>,<%=adsrs("id")%>)">
	                <a title="<%=adsrs("intro")%>" href="<%=adsrs("url")%>" target="<%=ttarg%>"><img border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>" /></a> 
				    </span>
				 <%end if%>
	<%          Case "swf"%>
	                <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http:/download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"; <%=GaoAndKuan%>><param name=movie value="<%=adsrs("gif_url")%>"><param name=quality value=high>
	  <%          Case "dai"%><%=adsrs("intro")%>
	  <embed src="<%=adsrs("gif_url")%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed></object>
	<%          Case else%>
	             <%If isend Then%>
	             <a title="<%=adsrs("intro")%>" href="javascript:;" onclick="alert('对不起，该广告今天已达到点击上限！');"><img border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>" /></a>
				 <%else%>
	             <a title="<%=adsrs("intro")%>" href="<%=adsrs("url")%>" target="<%=ttarg%>"><img border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>" /></a>
				 <%end if%>
	<%
	 End Select%>
	 <%
	adsrs.close
	set adsrs=nothing
	Conn.close
	set Conn=nothing 
	%>
	 </body>
	</html>
<%
 End Sub
 
 
'记录点击广告
Sub HitsGuangGao()
dim Url,getid,getclick,geturl,adssql,RSObj,SqlStr,getip,sitename,allowclick,place
		getid=KS.ChkClng(KS.S("id"))
		set RSObj=server.createobject("adodb.recordset")
		adssql="Select top 1 id,url,click,sitename,allowclick,place from KS_Advertise where id="&getid
		RSObj.open adssql,Conn,1,3
		if (rsobj.eof and rsobj.bof) then
		 rsobj.close
		 set rsobj=nothing
		 exit sub
		end if
		getclick=RSObj(2)+1
		sitename=RSOBJ(3)
		RSObj(2)=getclick
		RSObj.Update
		Url=RSObj(1)
		allowclick=RSObj("allowclick")
		place=RSObj("place")
		RSObj.Close
		If allowclick>0 Then '判断每天限制次数
		   Dim CurrClickNum:CurrClickNum=KS.ChkClng(Conn.Execute("Select count(1) From KS_Adiplist Where adid=" & GetId & " and DateDiff(" & DataPart_D & ",time," & SQLNowString &")<1")(0))
		   If CurrClickNum >= allowclick Then
		     KS.Die ""
		   End If
		End If
		
		'暂且关闭记录IP功能
		SqlStr="select top 1 * from KS_Adiplist where 1=0"
		RSObj.open SqlStr,Conn,1,3
		RSObj.AddNew
		RSObj("adid") = getid
		RSObj("time") = now()
		RSObj("ip") = KS.GetIP
		RSObj("class") = 2
		RSObj.update
		RSObj.close
		set RSObj=nothing 
		
		'========点广告加积分==================
		 if KS.Setting(166)="1" And KS.ChkClng(KS.Setting(167))>0 Then
		   If KS.C("UserName")<>"" Then
		      getid=KS.ChkClng(right(year(now),2)& "" & month(now) & "" & day(now)) & "" & getid  '每天产生不同的ID号，以便第二天增加积分
			  
			  If KS.Setting(145)="0" Then  '积分
					If Conn.Execute("Select top 1 * From KS_LogScore Where UserName='" & KS.C("UserName") & "' and year(adddate)=year(" & SQLNowString  &") and month(adddate)=month(" & SQLNowString &") and day(adddate)=day(" & SQLNowString & ") and channelid=1000 and infoid=" & getid).Eof Then
			          Call  KS.ScoreInOrOut(KS.C("UserName"),1,KS.ChkClng(KS.Setting(167)),"系统","点击广告[" & sitename & "(" & url & ")]所得!!",1000,getid)
			        End If
			  ElseIf KS.Setting(145)="1" Then '点券
			        If Conn.Execute("Select top 1 * From KS_LogPoint Where UserName='" & KS.C("UserName") & "' and year(adddate)=year(" & SQLNowString  &") and month(adddate)=month(" & SQLNowString &") and day(adddate)=day(" & SQLNowString & ") and channelid=1000 and infoid=" & getid).Eof Then
					  Call KS.PointInOrOut(1000,getid,KS.C("UserName"),1,KS.ChkClng(KS.Setting(167)),"系统","点击广告[" & sitename & "(" & url & ")]所得!!",0)
			        End If
			  Else
			       If Conn.Execute("Select top 1 * From KS_LogMoney Where UserName='" & KS.C("UserName") & "' and year(PayTime)=year(" & SQLNowString  &") and month(PayTime)=month(" & SQLNowString &") and day(PayTime)=day(" & SQLNowString & ") and channelid=1000 and infoid=" & getid).Eof Then
			          Call KS.MoneyInOrOut(KS.C("UserName"),KS.C("UserName"),KS.ChkClng(KS.Setting(167)),4,1,now,0,"System","点击广告[" & sitename & "(" & url & ")]所得!!",1000,getid,0)
			        End If
			  
					    
			  End If
			  
			  
			  
			  
			  
			  
		   End If
		 End If
		'=====================================
End Sub
 
End Class
 %>  

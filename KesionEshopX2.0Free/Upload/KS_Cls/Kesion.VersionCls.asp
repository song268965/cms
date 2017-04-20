<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Const ChannelNotOnStr="4,6,7,8,9,10"   '定义关闭的模块,请不要随便更改

'获得当前版本号
Function GetVersion()
  GetVersion=GetV("../../")
End Function
Function GetVer()
  GetVer=GetV("../")
End Function
Function GetV(dir)
    Dim MyVerSion
	Dim Doc:set Doc = CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	Doc.async = false
	Doc.setProperty "ServerHTTPRequest", true 
	Doc.load(Server.MapPath(dir &"config/version.xml"))
	if Doc.readystate=4 and Doc.parseError.errorCode=0 Then 
		Dim Node:Set Node= Doc.documentElement.selectSingleNode("//kesioncms/version")
		If Not Node Is Nothing Then MyVerSion=Node.text
	Else
	  Doc.load(Server.MapPath("/config/version.xml"))
	  if Doc.readystate=4 and Doc.parseError.errorCode=0 Then 
	     Set Node= Doc.documentElement.selectSingleNode("//kesioncms/version")
		If Not Node Is Nothing Then MyVerSion=Node.text
	  End If
	end if
    If MyVersion="" Then MyVerSion="2.0"
	GetV=MyVersion
End Function




Class KesionCls
	  Private Sub Class_Initialize()
      End Sub
	  Private Sub Class_Terminate()
	  End Sub
	 
	  '系统版本号
	  Public Property Get KSVer
		KSVer="KesionEshop X" & GetV("../") &" 商城版 Free"
	  End Property 
	  
	  '系统缓存名称,如果你的一个站点下安装多套科汛系统，请分别将各个目录下的系统的缓存名称设置成不同
	  Public Property Get SiteSN
	    If cbool(EnabledSubDomain) Then '如果启用二级域名，则SiteSN必须用固定值
		  SiteSN="KS_X2.0"
		Else
		  SiteSN="KS_X2.0" & Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME")), "/", ""), ".", "")  
	    End If
	  End Property
	   
End Class
%>
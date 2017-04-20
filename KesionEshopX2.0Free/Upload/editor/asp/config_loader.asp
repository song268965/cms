<!--#include file="../../conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->

<%
    ''==============KESION 修改==========================
    dim ks:set ks=new publiccls
	dim ksuser:set ksuser=new usercls
	dim IsLogin:IsLogin=KSUser.UserLoginChecked
    If IsLogin=false then ks.die "nologin"   '启用该句 google chrome 等浏览器无法使用抓图功能
    '==============KESION 修改==========================


	Set json = new ASPJson
    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    Set stream = Server.CreateObject("ADODB.Stream")

    stream.Open()
    stream.Charset = "gbk"
    stream.LoadFromFile Server.MapPath( "config.json" )

    content = stream.ReadText()

    Set commentPattern = new RegExp
    commentPattern.Multiline = true
    commentPattern.Pattern = "/\*[\s\S]+?\*/"
    commentPattern.Global = true
    content = commentPattern.Replace(content, "")
    json.loadJSON( content )

    Set config = json.data
%>
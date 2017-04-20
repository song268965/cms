<!--#include file="ASPJson.class.asp"-->
<!--#include file="config_loader.asp"-->
<!--#include file="Uploader.class.asp"-->
<%
    uploadTemplateName = Session.Value("ueditor_asp_uploadTemplateName")

    Set up = new Uploader
    up.MaxSize = config.Item( uploadTemplateName & "MaxSize" )
    up.FileField = config.Item( uploadTemplateName & "FieldName" )
    up.PathFormat = config.Item( uploadTemplateName & "PathFormat" )
	
	
		
	'==============KESION 修改==========================
	Dim allowPaths
	If ksuser.groupid<>1 then '前台会员限制只能选择自己上传的文件
      allowPaths=KS.ReturnChannelUserUpFilesDir(0,KSUser.GetUserInfo("userid"))
	Else
	  allowPaths=KS.GetUpFilesDir()
	End If
	if (right(allowPaths,1)<>"/") then allowPaths=allowPaths &"/"
   up.PathFormat = allowPaths & config.Item( uploadTemplateName & "PathFormat" )
	'==============KESION 修改==========================

	

    If Not IsEmpty( Session.Value("base64Upload") ) Then
        up.UploadBase64( Session.Value("base64Upload") )
    Else
        up.AllowType = config.Item( uploadTemplateName & "AllowFiles" )
        up.UploadForm()
    End If

    Set json = new ASPJson

    With json.data
        .Add "url", up.FilePath
        .Add "original", up.OriginalFileName
        .Add "state", up.State
        .Add "title", up.OriginalFileName
    End With
    
    json.PrintJson()
%>
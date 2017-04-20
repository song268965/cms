<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

'编辑器类型
Function GetEditorType()
 GetEditorType="Baidu"     '取值Baidu或CKEditor
End Function

'百度编辑器类型 取值UE或是UM
Function GetEditorTag()
   GetEditorTag="UE"     '取值 UE、UM
End Function



'输出编辑器头文件
Function EchoUeditorHead()
    Dim str,KS
	Set KS=New PublicCls
	If GetEditorType()="CKEditor" Then
		str=str & "<script type=""text/javascript"" src=""" & KS.Setting(3) & "ckeditor/ckeditor.js""></script>" &vbcrlf
	ElseIf GetEditorTag()="UM" Then
		str="<link href=""" & KS.Setting(3) & "umeditor/themes/default/css/umeditor.css"" type=""text/css"" rel=""stylesheet"">"
		str=str & "<script type=""text/javascript"" src=""" & KS.Setting(3) & "umeditor/third-party/jquery.min.js""></script>" &vbcrlf
		str=str & "<script type=""text/javascript"" charset=""utf-8"" src=""" & KS.Setting(3) & "umeditor/umeditor.config.js""></script>" &vbcrlf
		str=str & "<script type=""text/javascript"" charset=""utf-8"" src=""" & KS.Setting(3) & "umeditor/umeditor.min.js""></script>" &vbcrlf
		str=str & "<script type=""text/javascript"" src=""" & KS.Setting(3) & "umeditor/lang/zh-cn/zh-cn.js""></script>" &vbcrlf
	Else
		str="<script type=""text/javascript"" charset=""utf-8"" src=""" & KS.Setting(3) & "editor/ueditor.config.js""></script>" &vbcrlf
		str=str & "<script type=""text/javascript"" charset=""utf-8"" src=""" & KS.Setting(3) & "editor/ueditor.all.js""> </script>"&vbcrlf
		str=str & "<script type=""text/javascript"" charset=""utf-8"" src=""" & KS.Setting(3) & "editor/lang/zh-cn/zh-cn.js""></script>"&vbcrlf
    End If
	EchoUeditorHead="<script>var installDir='" & KS.Setting(3) &"';</script>" & vbcrlf & Str
	Set KS=Nothing
End Function

'输出编辑器
Function EchoEditor(FieldName,DefaultValue,ToolBar,Width,Height)
 Dim str
 If GetEditorType()="CKEditor" Then
	 str= "<textarea id=""" & fieldname &""" name=""" & fieldname &""">"& Server.HTMLEncode(DefaultValue) &"</textarea><script type=""text/javascript"">CKEDITOR.replace('" & fieldname &"', {width:""" & Width &""",height:""" & height & """,toolbar:""" & ToolBar & """,filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script>"
 Else
	 str= "<script id=""" & FieldName & """ name=""" & FieldName & """ type=""text/plain"" style=""width:" & Width & ";height:" & Height & ";"">" &DefaultValue&"</script>"
	 str= str & "<script>setTimeout(""baidu" & FieldName & " = " & GetEditorTag() &".getEditor('" & FieldName &"',{toolbars:[" & GetEditorToolBar(ToolBar) &"],wordCount:false,autoHeightEnabled:false,scaleEnabled:false,minFrameHeight:420 });"",10);</script>"
 End If
 EchoEditor=str
End Function
'判断编辑器有没有内容
Function GetEditorContent(FieldName)
 Dim str
 If GetEditorType()<>"CKEditor" Then
   str="baidu" & FieldName & ".hasContents()"
 Else
   str="CKEDITOR.instances."& FieldName &".getData()"
 End If
   GetEditorContent=str
End Function
'编辑器得到焦点
Function GetEditorFocus(FieldName)
 Dim str
 If GetEditorType()<>"CKEditor" Then
   str="baidu" & FieldName & ".focus();"
 Else
   str="CKEDITOR.instances."& FieldName &".focus();"
 End If
   GetEditorFocus=str
End Function
'向编辑器插入内容
Function InsertEditor(FieldName,codestr)
 Dim str
 If GetEditorType()<>"CKEditor" Then
  str="baidu"& FieldName &".execCommand('insertHtml', " & codestr & ");"
 Else
  str="CKEDITOR.instances."& FieldName &".insertHtml(" &codestr & ");"
 End If
  InsertEditor=str
End Function
'编辑器设置初始值
Function EditorSetContent(FieldName,Content)
 Dim str
 If GetEditorType()<>"CKEditor" Then
  str="baidu"& FieldName &".setContent('" &Content & "');"
 Else
  str="CKEDITOR.instances."& FieldName &".setData('" &Content & "');"
 End If
  EditorSetContent=str
End Function


	
'百度编辑器工具栏目定义
Function GetEditorToolBar(TypeFlag)
   Dim Str
   SELECT Case Lcase(TypeFlag)
     Case "basic"
	   Str="['fullscreen', 'source', '|', 'undo', 'redo', '|','bold', 'italic', 'underline', 'fontborder', 'strikethrough', 'superscript', 'subscript', 'removeformat','|','insertimage', 'emotion']"
     Case "nosourcebasic"
	   Str="['fullscreen',  'undo', 'redo', '|','bold', 'italic', 'underline', 'fontborder', 'strikethrough', 'superscript', 'subscript', 'removeformat','|','insertimage', 'emotion']"
	 Case "newstool"
	   Str="['fullscreen', 'source', '|', 'undo', 'redo', '|','bold', 'italic', 'underline', 'fontborder', 'strikethrough', 'superscript', 'subscript', 'removeformat', 'formatmatch', 'autotypeset', 'blockquote', 'pasteplain', '|', 'forecolor', 'backcolor', 'insertorderedlist', 'insertunorderedlist', 'selectall', 'cleardoc', '|','rowspacingtop', 'rowspacingbottom', 'lineheight', '|', 'customstyle', 'paragraph', 'fontfamily', 'fontsize', '|','directionalityltr', 'directionalityrtl','indent', '|', 'justifyleft', 'justifycenter', 'justifyright', 'justifyjustify', '|', 'touppercase', 'tolowercase', '|','link', 'unlink', 'anchor', '|', 'imagenone', 'imageleft', 'imageright', 'imagecenter', '|', 'insertimage', 'attachment','emotion', 'scrawl', 'insertvideo', 'music',  'map', 'gmap', 'insertframe','insertcode', 'webapp', 'template', 'background', '|', 'horizontal', 'date', 'time', 'spechars', 'snapscreen', 'wordimage', '|','inserttable', 'deletetable', 'insertparagraphbeforetable', 'insertrow', 'deleterow', 'insertcol', 'deletecol', 'mergecells', 'mergeright', 'mergedown', 'splittocells', 'splittorows', 'splittocols', 'charts', '|', 'pagebreak','preview', 'searchreplace', 'help', 'drafts']"
	 Case Else
	   Str="['fullscreen', 'source', '|', 'undo', 'redo', '|','bold', 'italic', 'underline', 'fontborder', 'strikethrough', 'superscript', 'subscript', 'removeformat', 'formatmatch', 'autotypeset', 'blockquote', 'pasteplain', '|', 'forecolor', 'backcolor', 'insertorderedlist', 'insertunorderedlist', 'selectall', 'cleardoc', '|','rowspacingtop', 'rowspacingbottom', 'lineheight', '|', 'customstyle', 'paragraph', 'fontfamily', 'fontsize', '|','directionalityltr', 'directionalityrtl','indent', '|', 'justifyleft', 'justifycenter', 'justifyright', 'justifyjustify', '|', 'touppercase', 'tolowercase', '|','link', 'unlink', 'anchor', '|', 'imagenone', 'imageleft', 'imageright', 'imagecenter', '|', 'insertimage', 'emotion', 'scrawl', 'insertvideo', 'music', 'map', 'gmap', 'insertframe','insertcode', 'webapp', 'pagebreak', 'template', 'background', '|', 'horizontal', 'date', 'time', 'spechars', 'snapscreen', 'wordimage', '|','inserttable', 'deletetable', 'insertparagraphbeforetable', 'insertrow', 'deleterow', 'insertcol', 'deletecol', 'mergecells', 'mergeright', 'mergedown', 'splittocells', 'splittorows', 'splittocols', 'charts', '|','print', 'preview', 'searchreplace', 'help', 'drafts']"
   End Select
     GetEditorToolBar=Str
End Function		
%>

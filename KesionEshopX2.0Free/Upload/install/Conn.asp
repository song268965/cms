<%
'Response.Buffer=True
Dim SqlNowString
Dim Conn,DBPath,ConnStr,Conn_Old

Sub OpenOldConn()
    On Error Resume Next
    ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../ks_data/kesioncmsx20.mdb")
    Set Conn_Old = Server.CreateObject("ADODB.Connection")
    Conn_Old.open ConnStr
    If Err Then 
	Err.Clear:Set Conn_Old = Nothing
	Response.Write "<script>KesionJS.Alert('数据库连接出错,检查/ks_data/kesioncmsx20.mdb是否存在!');</script>"  
	Response.End
	end if
End Sub
Sub CloseConn()
    On Error Resume Next
	Conn.close:Set Conn=nothing
	conn_old.close:set conn_old=nothing
End sub

Sub OpenConn_install(C_Str,D_Type)
    On Error Resume Next
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.open C_Str
    If Err Then 
	 Err.Clear:Set conn = Nothing
	Response.Write "<script>KesionJS.Alert('数据库连接出错," & Err.Description &"请检查数据库是否存在!','');</script>"  
	Response.End
    end if
End Sub

%>

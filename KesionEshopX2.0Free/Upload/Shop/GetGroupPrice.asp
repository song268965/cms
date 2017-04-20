<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Dim KS,t,l
Set KS=New PublicCls
Set KSUser = New UserCls
Dim ProID,Tips,RS,SqlStr,RealPrice
ProID = KS.ChkClng(KS.S("ProID"))
T     = KS.S("T")
l     = KS.ChkClng(KS.S("l"))
 If ProID = 0 Or Cbool(KSUser.UserLoginChecked)=false Then
    IF T="p" Then
	   If l=1 Then
        Tips = "<a href=""javascript:ShowLogin();"">登录</a>"
	   Else
        Tips = "登录查看您的价格。<a href=""javascript:ShowLogin();""><font color=red>登录</font></a> <a href=""" & KS.Setting(3) & "user/reg""><font color=red>注册</font></a>"
	   End If
	Else
        Tips = "登录查看可获得积分。<a href=""javascript:ShowLogin();""><font color=red>登录</font></a> <a href=""" & KS.Setting(3) & "user/reg""><font color=red>注册</font></a>"
	End If
 Else
        Dim Discount:Discount=KS.U_S(KSUser.GroupID,17)
		Dim JFDiscount:JFDiscount=KS.U_S(KSUser.GroupID,18)
		If Not IsNumeric(Discount) Then Discount=0
		If Not IsNumeric(JFDiscount) Then JFDiscount=0
        SqlStr = "SELECT Top 1 Price_Member,isdiscount,islimitbuy,limitbuyprice,istype,score,VipPrice FROM KS_Product Where ID=" & ProID
        Set RS = Server.CreateObject("ADODB.Recordset")
        RS.Open SqlStr, conn, 1, 1
        If Not RS.bof  Then
		  if rs("islimitbuy")<>0 then
					Tips = "￥" & KS.GetPrice(rs(3)) & "元(抢购价)"
					RealPrice=KS.GetPrice(rs(3))
					IF T<>"p" Then
					Tips = "0"
					end if
		  else
			  If Discount=0 or KS.ChkClng(rs("isdiscount"))=0 Then
					Tips = "￥" & KS.GetPrice(rs(0)) & "元"
					if rs("istype")<>"0" then tips=rs("score")&"分+" & tips
					RealPrice=KS.GetPrice(rs(0))
			  Else
					Tips = "￥" & KS.GetPrice(rs(0)*discount/10) & "元<span>(" & Discount & "折)</span>"  
					if rs("istype")<>"0" then tips=rs("score")&"分+" & tips
					RealPrice=KS.GetPrice(rs(0)*discount/10)
			  End If
			  
			  If KS.U_S(KSUser.GroupID,21)="1" and rs("vipprice")<>"0" then
			   Tips = "￥" & KS.GetPrice(rs("vipprice")) & "元"
			   RealPrice=KS.GetPrice(rs("vipprice"))
			  end if	
			     
			  IF T<>"p" Then
				   If JFDiscount=0 Then
					Tips = "<font color=""red"">0</font>"
				   ElseIf JFDiscount=1 Then
					Tips = "<font color=""red"">" & KS.ChkClng(RealPrice) & "</font>分"
				   Else
				    If l=1 Then
					Tips = "<font color=""red"">" & KS.ChkClng(RealPrice*JFDiscount) & "</font>分<span>"
					else
						if (RealPrice*JFDiscount)<1 then
							Tips = "<font color=""red"">0" & RealPrice*JFDiscount & "</font>分<span>(实际价格的" & JFDiscount & "倍积分)</span>"'bug改动
						else
							Tips = "<font color=""red"">" & RealPrice*JFDiscount & "</font>分<span>(实际价格的" & JFDiscount & "倍积分)</span>"'bug改动
						end if
					end if
				   End If
			  End If
		  end if
        End If
        rs.Close
        Set rs = Nothing
End If
Response.Write "document.write('" & Tips & "');"
If EnabledSubDomain Then
 response.write "document.domain=""" & RootDomain &""";" &vbcrlf
end if

Call CloseConn()
Set KS=Nothing	
%>

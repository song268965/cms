<!--#include file="../conn.asp"-->
<!--#include file="../ks_cls/kesion.commoncls.asp"-->
<!--#include file="PackCart.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Dim KS:Set KS=New PublicCls
Dim Num:Num=Conn.Execute("Select count(1) From KS_Shoppingcart Where flag=0 and UserName='"& GetUserID & "'")(0)
Response.Write "document.writeln(""<img src='" & KS.GetDomain & "Images/user/log/5.gif'> <a href='" & KS.GetDomain & "shop/shoppingcart.asp' target='_blank'>购物车 <span style='color:#ff0000'>" & Num & "</span> 件商品</a> | <a href='" & KS.GetDomain & "user/user_Favorite.asp' target='_blank'>收藏夹</a> | <a href='" & KS.GetDomain & "user/user_order.asp' target='_blank'>我的订单</a>"");"
Set KS=Nothing
CloseConn
%>

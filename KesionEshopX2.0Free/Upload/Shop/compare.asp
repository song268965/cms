<!--#include file="../conn.asp"-->
<!--#include file="../ks_cls/kesion.commoncls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************


Dim ShowFieldStr:ShowFieldStr="商品名称|title@商品图片|photourl@市场价|price@商城价|price_member@商品商标|trademarkname@添加时间|adddate@推荐等级|rank@商品单位|unit@点击数|hits@商品录入|inputer"
Dim KS:Set KS=New PublicCls
Dim IDS,RS,SQLStr,XML,ShowFieldArr,FieldArr,I,K,Templates,Node
Ids=Request("Ids")
If KS.IsNul(Ids) Then
  KS.Die "<script>alert('非法参数!');window.close();</script>"
End If 
Ids=KS.FilterIds(Ids)
If IDS="" Then
  KS.Die "<script>alert('非法参数!');window.close();</script>"
End If
Set RS=Server.CreateObject("ADODB.RECORDSET")
RS.Open "Select FieldName,Title From KS_Field Where ChannelID=5 and fieldtype<>0 Order By OrderID Asc",conn,1,1
Do While Not RS.Eof
  ShowFieldStr=ShowFieldStr & "@" & rs(1) & "|" & lcase(rs(0))
  RS.MoveNext
Loop
RS.Close

SqlStr="Select * From KS_Product Where Id In (" & Ids & ") order by id desc"
RS.Open SqlStr,conn,1,1
If RS.Eof And RS.Bof Then
 RS.Close: Set RS=Nothing
 KS.Die "<script>alert('找不到商品!');window.close();</script>"
End If
Set XML=KS.RsToXml(RS,"row","root")
If Not IsObject(Xml) Then
 KS.Die "<script>alert('找不到商品!');window.close();</script>"
End If
ShowFieldArr=Split(ShowFieldStr,"@")
echo "<table width='100%' cellspacing='0' class='compare' cellpadding='0' border='0'>" &vbcrlf
For I=0 To Ubound(ShowFieldArr)
   FieldArr=Split(ShowFieldArr(i),"|")
   echo "<tr>"
   echo "<td align='right' class='title' width='100'>" & FieldArr(0) & "：</td>"&vbcrlf
    For K=1 To XML.DocumentElement.SelectNodes("row").length
	  Set Node=XML.DocumentElement.SelectNodes("row").item(k-1)
	  echo "<td class='c" & K & " pro_pic'>" &vbcrlf
	  if FieldArr(1)="photourl" then 
	   echo "<img src='" & Node.selectsinglenode("@photourl").text & "' width='150' height='113'/>"
	  elseif FieldArr(1)="tid" then 
	   echo KS.GetClassNP(Node.selectsinglenode("@tid").text)
	  elseif FieldArr(1)="inputer" then
	   echo "<a href='../space/?" & Node.selectsinglenode("@inputer").text & "' target='_blank'>" & Node.selectsinglenode("@inputer").text & "</a>"
	  elseif FieldArr(1)="title" Then
	   echo "<a href='" & KS.GetItemURL(5,Node.selectsinglenode("@tid").text,Node.selectsinglenode("@id").text,Node.selectsinglenode("@fname").text,Node.selectsinglenode("@adddate").text) & "' target='_blank'>" & Node.selectsinglenode("@title").text & "</a>"
	  else
	   echo Node.selectsinglenode("@" &FieldArr(1)).text
	  end if
	  echo "</td>"&vbcrlf
	Next
   echo "</tr>"&vbcrlf
Next
			  echo "<tr>" & vbcrlf
			  echo "<td align='right' width='100' class='title'>&nbsp;</td>"  &vbcrlf
				For K=1 To XML.DocumentElement.SelectNodes("row").length
			       echo "<td class='c" & K & "' style='text-align:center'><a href='javascript:;' onclick="" $('.compare').find('.c" & k & "').html('');"" class='qx'>取消对比</a></td>" 
				Next
			   echo "</tr>" & vbcrlf

echo "</table>"

Dim KSR:Set KSR=New Refresh

 Dim TpDir:TpDir=KS.SSetting(63)
 If KS.IsNUL(TpDir) Then TpDir=KS.Setting(3) & KS.Setting(90) & "空间模板/企业通用/product_compare.html"
 Dim Template:Template = KSR.LoadTemplate(TpDir)
			 
Template=Replace(template,"{$ShowCompareList}",templates)
Template=KSR.KSLabelReplaceAll(Template)
Set KSR=Nothing

KS.Echo Template
Set KS=Nothing
CloseConn

Function Echo(str)
  Templates=Templates & str &vbcrlf
End Function
%>
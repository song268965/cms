<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Template.asp"-->
<!--#include file="../KS_Cls/Kesion.IFCls.asp"-->
<%
'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************

Dim KSCls
Set KSCls = New GroupBuyShow
KSCls.Kesion()
Set KSCls = Nothing

Class GroupBuyShow
        Private KS, KSR,Product,LoginTf,hasbuynum
		Private GroupBuy,K,ID
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		  Call showmain()
		End Sub
		%>
		<!--#include file="../KS_Cls/Kesion.IFCls.asp"-->
		<%
		Sub ShowMain()
			 Dim FileContent
			 FileContent = KSR.LoadTemplate(KS.Setting(138))    
			 FCls.RefreshType = "groupbyIndex" '设置刷新类型，以便取得当前位置导航等
			 FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
			 GetQueryParam
			 LoadGroupBuyList
			 hasbuynum=KS.ChkClng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.ordertype=1 and i.proid=" & GroupBuy(0,0))(0))
			 HasBuyNum=HasBuyNum+KS.ChkClng(GroupBuy(18,0))  '加上作弊的件数
			 Immediate=false
			 Scan FileContent
			 Templates=KSR.KSLabelReplaceAll(Templates)
			 Response.write RexHtml_IF(Templates)
		End Sub
		

		
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			   case "groupbuy"
			         Select case lcase(sTokenName)
					  case "todaygroupbuylink"  
					   If KS.ChkClng(KS.Setting(179))=1 Then Echo KS.GetDomain & "groupbuy/" Else Echo KS.GetDomain & "shop/groupbuy.asp"
					  case "historygroupbuylink"  
					   If KS.ChkClng(KS.Setting(179))=1 Then Echo KS.GetDomain & "groupbuy/history/" Else Echo KS.GetDomain & "shop/groupbuy.asp?flag=history"
					 End Select
			   case "product"
					 Select  case lcase(sTokenName)
					  case "id" echo GroupBuy(0,0)
					  case "writecomment" If GroupBuy(20,0)<>"0" then echo "<script tyle=""text/Javascript"" src=""" & KS.GetDomain & "plus/digmood/Comment.asp?Action=Write&ChannelID=1000&InfoID=" & GroupBuy(0,0) & """></script>"
					  case "showcomment1" If GroupBuy(20,0)<>"0" then echo "<script src=""" & KS.GetDomain & "ks_inc/Comment.page.js"" type=""text/javascript""></script><script type=""text/javascript"" defer>var from3g=0;Page(1,1000,'" & GroupBuy(0,0) & "','Show','"& ks.GetDomain & "');</script><div id=""c_" & GroupBuy(0,0) & """></div><div id=""p_" & GroupBuy(0,0) & """ align=""right""></div>"
					  case "showcomment" 
					    If GroupBuy(20,0)<>"0" then
						 echo "<script src=""" & KS.GetDomain & "ks_inc/Comment.page.js""></script><script defer>var from3g=0;Page(1,1000,'" & GroupBuy(0,0) & "','Show',5,'"& ks.GetDomain & "');</script><div id=""c_" & GroupBuy(0,0) & """></div><div id=""p_" & GroupBuy(0,0) & """ align=""right""></div>"
						end if
					  case "subject" echo GroupBuy(1,0)
					  case "adddate" echo GroupBuy(3,0)
					  case "intro" echo GroupBuy(4,0)
					  case "cartlink" If KS.ChkClng(KS.Setting(179))=1 Then Echo KS.GetDomain & "groupbuy/cart-" &GroupBuy(0,0) & ".html"  Else Echo KS.GetDomain & "shop/groupbuycart.asp?id=" & GroupBuy(0,0)
					  case "photourl" if KS.IsNul(GroupBuy(5,0)) Then Echo "../images/nopic.gif" Else Echo GroupBuy(5,0)
					  case "bigphoto" if KS.IsNul(GroupBuy(19,0)) Then Echo "../images/nopic.gif" Else Echo GroupBuy(19,0)
					  case "highlights" echo GroupBuy(6,0)
					  case "protection" echo GroupBuy(7,0)
					  case "notes" echo GroupBuy(8,0)
					  case "enddate" echo groupbuy(2,0)
					  case "timetips"
					    if GroupBuy(15,0)=1 then
						 echo "<span style='color:#999999'>本团购已锁定</span>"
					    Elseif GroupBuy(16,0)=1 then
						 echo "<span style='color:#999999'>本团购已结束</span>"
						ElseIF DateDiff("s",now(),GroupBuy(3,0))>0 Then
						 echo "<span style=""font-weight:bold;color:green"">本次团购未开始，开始时间：" & GroupBuy(3,0) & "</span>"
						Elseif DateDiff("s",now(),groupbuy(2,0))<0 then
						 echo "<strong>本次团购结束于<br/> " & groupbuy(2,0) & "</strong>"
						else
						 echo "距离团购结束还有：<br/><script src=""" & KS.GetDomain &"shop/js/lefttimes.js""></script>"
						 echo "<ul id=""counter" & GroupBuy(0,0) & """ style=""font-weight:bold""></ul>"
						 echo "<script type='text/javascript'>showtime('counter" & GroupBuy(0,0) & "'," & DateDiff("s",now(),groupbuy(2,0)) & ");</script>"
						End If
			  
					  case "endsecond" echo DateDiff("s",now(),groupbuy(2,0))
					  case "minnum" If KS.ChkClng(GroupBuy(11,0))=0 Then echo "不限制" else echo GroupBuy(11,0)&"人"
					  case "price_original" echo groupbuy(12,0)
					  case "discount" echo groupbuy(13,0)
					  case "price" echo groupbuy(14,0)
					  case "pricesave" echo groupbuy(12,0)-groupbuy(14,0)
					  case "hasbuynum" echo hasbuynum
					  case "userlist"
					    Dim RSU:Set RSU=Conn.Execute("select top 1000 o.contactman,o.mobile from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.ordertype=1 and i.proid=" & GroupBuy(0,0) & " order by i.id desc")
						If Not RSU.Eof Then
						   Do While NOt RSU.Eof
						    echo "<li>" & KS.Gottopic(rsu(0),2)&"**   手机：" & left(rsu(1)&"000",3) & "*****" & right("000"&rsu(1),3) &"</li>"
						    RSU.MoveNext
						   Loop
						Else
						    echo "<li>此单还没有用户购买</li>"
						End If
						RSU.Close:Set RSU=Nothing
					   case "buytips"
					     if DateDiff("s",now(),groupbuy(2,0))<0 then
						      echo "本团购已结束"
						 Elseif KS.ChkClng(hasbuynum) < KS.ChkClng(GroupBuy(11,0)) Then
						      dim wid:wid=KS.ChkClng(KS.ChkClng(hasbuynum)*200/KS.ChkClng(GroupBuy(11,0)))
						      echo "<div class=""tipometer"">" &vbcrlf
                              echo " <div class=""tipping_point"" style=""left:" & wid & "px""></div>"&vbcrlf
							  echo "  <div class=""progress_bar"">"&vbcrlf
							  echo "      <div class=""post_tipped"" style=""width:" & wid & "px""></div>"&vbcrlf
							  echo "  </div>"&vbcrlf
							  echo " <div class=""l min"">0</div>"&vbcrlf
							  echo " <div class=""r max"">" & GroupBuy(11,0) & "</div>"&vbcrlf
							  echo "<div class=""c""></div>"&vbcrlf
							  echo "</div>"&vbcrlf
							  echo "<div class=""done"" style=""text-align:center"">距离团购人数还差" & KS.ChkClng(GroupBuy(11,0))-KS.ChkClng(hasbuynum) & "人。</div>"
						 Else
						      echo "<div class=""tgpeople""><img src=""" & KS.GetDomain & "shop/images/suc.gif"" align=""absmiddle"" width=27 height=28/> 已成团，可继续购买...</div>"
							  echo "<div class=""tgpeople"">"
							  if IsDate(GroupBuy(17,0)) Then echo hour(GroupBuy(17,0)) & "点" & minute(GroupBuy(17,0)) & "分"
							  echo "达到最低团购人数：" & GroupBuy(11,0) & "人"
							  echo "</div>"
						 End If
					 end select
		    End Select 
        End Sub 
		
		
		Sub GetQueryParam()
		  ID=KS.ChkClng(KS.S("ID"))
		  If ID=0 Then
		   Call KS.ShowTips("error","参数出错!")
		   Response.End
		  End If
		End Sub
		


		
		Sub LoadGroupBuyList()
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select TOP 1 ID,Subject,ActiveDate,adddate,intro,photourl,Highlights,Protection,Notes,AllowBMFlag,AllowArrGroupID,minnum,Price_Original,Discount,Price,locked,EndTF,MinnumTime,HasBuyNum,bigphoto,Comment,Changes,ChangesUrl From KS_GroupBuy Where ID=" & ID & " Order By Id Desc",Conn,1,1
		 If RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   Call KS.ShowTips("error","参数出错或本活动已关闭!")
		   Response.End()
		 Else
		   GroupBuy=RS.GetRows(-1)
		 End If
		 RS.Close:Set RS=Nothing
		 if GroupBuy(21,0)="1" then response.Redirect(GroupBuy(22,0))
		End Sub


End Class
%>

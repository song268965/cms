<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
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
Set KSCls = New GroupBuyIndex
KSCls.Kesion()
Set KSCls = Nothing

Class GroupBuyIndex
        Private KS, KSR,KSUser,Param,categoryid,PriceArr
		Private GroupBuy,K,CurrentPage,totalPut,MaxPerPage,hasbuynum
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  MaxPerPage=20         rem 定义每页显示条数
		  PriceArr=Array("所有|1=1","50元以下|price<50","50~100元|price>=50 and price<=100","100~200元|price>=100 and price<=200","200~300元|price>=200 and price<=300","300~500元|price>=300 and price<=500","500~1000元|price>=500 and price<=1000","1000元以上|price>1000") rem 定义按价格搜索
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
			 FileContent = KSR.LoadTemplate(KS.Setting(137))    
			 FCls.RefreshType = "groupbyIndex" '设置刷新类型，以便取得当前位置导航等
			 FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
			 GetQueryParam
			 FileContent=RexHtml_IF(FileContent)
			 LoadGroupBuyList
			 Immediate=false
			 Scan FileContent
			 Templates=KSR.KSLabelReplaceAll(Templates)
			 Response.write Templates
		End Sub
		
		
		Sub ParseArea(sTokenName, sTemplate)
			Select Case lcase(sTokenName)
			 case "groupbuylist"
			  If IsArray(GroupBuy) Then
			    hasbuynum=0
			    For K=0 To Ubound(GroupBuy,2)
				  hasbuynum=KS.ChkClng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.ordertype=1 and i.proid=" & GroupBuy(0,k))(0))
				  hasbuynum=hasbuynum+KS.ChkClng(GroupBuy(15,k))
				  Scan sTemplate
				Next
			  End If
			End Select 
        End Sub 
		
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			   case "groupbuy"
			         Select case lcase(sTokenName)
					  case "todaygroupbuylink"  
					   If KS.ChkClng(KS.Setting(179))=1 Then Echo KS.GetDomain & "groupbuy/" Else Echo KS.GetDomain & "shop/groupbuy.asp"
					  case "historygroupbuylink"  
					   If KS.ChkClng(KS.Setting(179))=1 Then Echo KS.GetDomain & "groupbuy/history/" Else Echo KS.GetDomain & "shop/groupbuy.asp?flag=history"
					  case "showcategory" call showcategory()
					  case "showprice" call showprice()
					  case "showpage"
					   echo KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false)
					 End Select
			   case"product"
					 Select  case lcase(sTokenName)
					  case "floor" echo k+1
					  case "id" echo GroupBuy(0,k)
					  case "linkurl" 
					   If GroupBuy(17,k)="1" then
					     Echo GroupBuy(18,k)
					   else
					     If KS.ChkClng(KS.Setting(179))=1 Then Echo KS.GetDomain & "groupbuy/show-" &GroupBuy(0,k) & ".html"  Else Echo KS.GetDomain & "shop/groupbuyshow.asp?id=" & GroupBuy(0,k)
					   end if
					  case "cartlink" 
					   If GroupBuy(17,k)="1" then
					     Echo GroupBuy(18,k)
					   else
					    If KS.ChkClng(KS.Setting(179))=1 Then Echo KS.GetDomain & "groupbuy/cart-" &GroupBuy(0,k) & ".html"  Else Echo KS.GetDomain & "shop/groupbuycart.asp?id=" & GroupBuy(0,k)
					   End If
					  case "title" echo GroupBuy(1,k)
					  case "price_original" echo GroupBuy(2,k)
					  case "price" echo GroupBuy(3,k)
					  case "discount" echo GroupBuy(4,k)
					  case "save" echo GroupBuy(2,k)-GroupBuy(3,k)
					  case "photourl"  if KS.IsNul(GroupBuy(5,k)) Then Echo KS.GetDomain & "images/nopic.gif" Else Echo GroupBuy(5,k)
					  case "bigphoto" if KS.IsNul(GroupBuy(16,k)) Then Echo KS.GetDomain & "images/nopic.gif" Else Echo GroupBuy(16,k)
					  case "note" echo GroupBuy(9,k)
					  case "adddate" echo year(GroupBuy(7,k)) & "年" &month(GroupBuy(7,k)) & "月"  & day(GroupBuy(7,k)) & "日"
					  case "hasbuynum" echo hasbuynum
					  case "limitbuynum" echo GroupBuy(11,k)
					  case "minnum" echo GroupBuy(10,k)
					  case "showclass" 
					   if (k+1) mod 2=0 then echo " class=""ma_r"""
					  case "timetips"
					    if GroupBuy(12,k)=1 then
						 echo "<span style='color:#999999'>本团购已锁定</span>"
					    Elseif GroupBuy(13,k)=1 then
						 echo "<span style='color:#999999'>本团购已结束</span>"
						ElseIF DateDiff("s",now(),GroupBuy(7,k))>0 Then
						 echo "<span style=""font-weight:bold;color:green"">本次团购未开始，开始时间：" & GroupBuy(7,k) & "</span>"
						Elseif DateDiff("s",now(),groupbuy(6,k))<0 then
						 echo "<strong>本次团购结束于<br/> " & groupbuy(6,k) & "</strong>"
						else
						 echo "<ul id=""counter" & GroupBuy(0,k) & """></ul>"
						 echo "<script type='text/javascript'>showtime('counter" & GroupBuy(0,k) & "'," & DateDiff("s",now(),groupbuy(6,k)) & ");</script>"
						End If
					 case "buytips"
					     if DateDiff("s",now(),groupbuy(6,k))<0 then
						      echo "本团购已结束"
						 Elseif KS.ChkClng(hasbuynum) < KS.ChkClng(GroupBuy(10,k)) Then
						      dim wid:wid=KS.ChkClng(KS.ChkClng(hasbuynum)*200/KS.ChkClng(GroupBuy(10,k)))
						      echo "<div class=""tipometer"">" &vbcrlf
                              echo " <div class=""tipping_point"" style=""left:" & wid & "px""></div>"&vbcrlf
							  echo "  <div class=""progress_bar"">"&vbcrlf
							  echo "      <div class=""post_tipped"" style=""width:" & wid & "px""></div>"&vbcrlf
							  echo "  </div>"&vbcrlf
							  echo " <div class=""l min"">0</div>"&vbcrlf
							  echo " <div class=""r max"">" & GroupBuy(10,k) & "</div>"&vbcrlf
							  echo "<div class=""c""></div>"&vbcrlf
							  echo "</div>"&vbcrlf
							  echo "<div class=""done"" style=""text-align:center"">距离团购人数还差" & KS.ChkClng(GroupBuy(10,k))-KS.ChkClng(hasbuynum) & "人。</div>"
						 Else
						      echo "<div class=""done""><img src=""" & KS.GetDomain & "shop/images/deal-buy-succ.gif"" align=""absmiddle"" width=27 height=28/> 团购已成功，还可以继续购买...</div>"
							  echo "<div style=""color:#666"">"
							  if IsDate(GroupBuy(14,k)) Then echo hour(GroupBuy(14,k)) & "点" & minute(GroupBuy(14,k)) & "分"
							  echo "达到最低团购人数：" & GroupBuy(10,k) & "人"
							  echo "</div>"
						 End If
					 end select
		    End Select 
        End Sub 
		
		sub showcategory()
			Dim RS:Set RS=Conn.Execute("select ID,CategoryName From KS_GroupBuyClass Order BY OrderID,ID")
			If NOT RS.Eof Then
				Dim BGStr:BGStr=""
				Dim ColorStr:ColorStr=""
				Dim LinkUrl:LinkUrl=""
				if categoryid=0 then
				  If KS.ChkClng(KS.Setting(179))=1 Then
				   echo "<li><a href="""  & KS.GetDomain & "groupbuy/history/"" style=""color:#fff;background:#CD1A01""><span>所有</span></a></li>"
				  Else
				   echo "<li><a href="""  & KS.GetDomain & "shop/groupbuy.asp?flag=history"" style=""color:#fff;background:#CD1A01""><span>所有</span></a></li>"
				  End If
				else
				 If KS.ChkClng(KS.Setting(179))=1 Then
				  echo "<li><a href="""  & KS.GetDomain & "groupbuy/history/""><span>所有</span></a></li>"
				 Else
				  echo "<li><a href="""  & KS.GetDomain & "shop/groupbuy.asp?flag=history""><span>所有</span></a></li>"
				 End If
				end if
				Do While NOT RS.Eof
						    If categoryid=RS("ID") Then  
							 BGStr=" style='color:#fff;background:#CD1A01'" 
							 ColorStr=" style='color:#fff'"
						    Else 
							 BGStr=""
							 ColorStr=""
							End If
							If KS.ChkClng(KS.Setting(179))=1 Then
							  LinkUrl=KS.GetDomain & "groupbuy/history.html?c=" & RS(0)
							Else
							  LinkUrl=KS.GetDomain & "shop/groupbuy.asp?flag=history&categoryid=" & RS(0)
							End If
						    echo "<li><a href=""" & LinkURL & """" & BGStr & "><span>" & RS(1) & "<em" & ColorStr &">(" & Conn.Execute("select count(1) from KS_GroupBuy Where Verific=1 and ClassID=" & RS(0))(0) & ")</em></span></a></li>"
				  RS.MoveNext
				Loop
			End If
			RS.Close:Set RS=Nothing
		end sub
		
		sub showprice()
		 dim i,pp
		 if categoryid<>0 then pp="&categoryid=" & categoryid
		 for i=0 to ubound(PriceArr)
		  If pp="" Then    '没有分类
		    If KS.ChkClng(KS.Setting(179))=1 Then
			  If KS.ChkClng(KS.S("P"))=i Then
			  echo "<li><a href=""" & KS.GetDomain & "groupbuy/history.html?p=" & i &""" style=""color:#fff;background:#CD1A01""><span>" & split(priceArr(i),"|")(0) & "</span></a></li>"
			  Else
			  echo "<li><a href=""" & KS.GetDomain & "groupbuy/history.html?p=" & i &"""><span>" & split(priceArr(i),"|")(0) & "</span></a></li>"
			  End If
			Else
			  If KS.ChkClng(KS.S("P"))=i Then
			  echo "<li><a href=""" & KS.GetDomain & "shop/groupbuy.asp?flag=history&p=" & i & pp &""" style=""color:#fff;background:#CD1A01""><span>" & split(priceArr(i),"|")(0) & "</span></a></li>"
			  Else
			  echo "<li><a href=""" & KS.GetDomain & "shop/groupbuy.asp?flag=history&p=" & i & pp & """><span>" & split(priceArr(i),"|")(0) & "</span></a></li>"
			  End If
			End If
		  Else
		    If KS.ChkClng(KS.Setting(179))=1 Then
			  If KS.ChkClng(KS.S("P"))=i Then
			  echo "<li><a href=""" & KS.GetDomain & "groupbuy/history.html?p=" & i &"&c=" & categoryid & """ style=""color:#fff;background:#CD1A01""><span>" & split(priceArr(i),"|")(0) & "</span></a></li>"
			  Else
			  echo "<li><a href=""" & KS.GetDomain & "groupbuy/history.html?p=" & i &"&c=" & categoryid & """><span>" & split(priceArr(i),"|")(0) & "</span></a></li>"
			  End If

			Else
			  If KS.ChkClng(KS.S("P"))=i Then
			  echo "<li><a href=""?flag=history&p=" & i & pp &""" style=""color:#fff;background:#CD1A01""><span>" & split(priceArr(i),"|")(0) & "</span></a></li>"
			  Else
			  echo "<li><a href=""?flag=history&p=" & i & pp & """><span>" & split(priceArr(i),"|")(0) & "</span></a></li>"
			  End If
			End If
		  End If
		 Next
		end sub
		
		Sub GetQueryParam()
		  If KS.S("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
		  Else
			  CurrentPage = 1
		  End If
		  If CurrentPage < 1 Then CurrentPage = 1
		   categoryid=KS.ChkClng(KS.S("categoryid"))
		End Sub

		
		Sub LoadGroupBuyList()
		 Param=" Where Endtf=0 and Verific=1 and Locked=0"
		 If categoryid<>0 Then Param=Param & " And ClassID=" & categoryid
		 If KS.ChkClng(KS.S("P"))<>0 and KS.ChkClng(KS.S("P"))<=Ubound(PriceArr) Then Param=Param & " and " & split(PriceArr(KS.ChkClng(KS.S("P"))),"|")(1)
		 Dim TopStr
		 If Request("flag")<>"history" Then 
		  TopStr=" top 9"
		  'Param=Param & " and datediff(" & DataPart_D & ",adddate," & SQLNowString & ")<=0"   
		  '如果想限制只调用本日的团购，请将上面语句前面的单引号去掉
		 Else
		   Param=Param & " and datediff(" & DataPart_D & ",adddate," & SQLNowString & ")>0"
		 End If
		 Dim SQLStr:SQLStr="Select  " & TopStr & " ID,Subject,Price_Original,Price,Discount,photourl,ActiveDate,AddDate,Intro,Notes,minnum,LimitBuyNum,Locked,EndTF,MinnumTime,HasBuyNum,bigphoto,Changes,ChangesUrl From KS_GroupBuy  " & Param & " Order By Id Desc"
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open SQLStr,Conn,1,1
		 If Not RS.Eof Then
		    If Request("flag")<>"history" Then 
			    GroupBuy=RS.GetRows(-1)
			Else
			        TotalPut= rs.recordcount
					If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
					End If
					GroupBuy=RS.GetRows(MaxPerPage)
			End If
		 End If
		 RS.Close:Set RS=Nothing
		End Sub
		
'伪静态分页
Public Function ShowPage()
		           Dim I, pageStr
				   pageStr= ("<div id=""fenye"" class=""fenye""><table border='0' align='right'><tr><td>")
					if (CurrentPage>1) then pageStr=PageStr & "<a href=""rating-" & channelid & "-" & infoid &"-" & projectid & "-" & CurrentPage-1 & ".html"" class=""prev"">上一页</a>"
				   if (CurrentPage<>PageNum) then pageStr=PageStr & "<a href=""rating-" & channelid & "-" & infoid & "-" & projectid & "-" & CurrentPage+1 & ".html"" class=""next"">下一页</a>"
				   pageStr=pageStr & "<a href=""rating-" & channelid &"-" & infoid & "-" & projectid & "-1.html"" class=""prev"">首 页</a>"
				 
					Dim startpage,n,j
					 if (CurrentPage>=7) then startpage=CurrentPage-5
					 if PageNum-CurrentPage<5 Then startpage=PageNum-10
					 If startpage<0 Then startpage=1
					 n=0
					 For J=startpage To PageNum
						If J= CurrentPage Then
						 PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
						Else
						 PageStr=PageStr & " <a class=""num"" href=""rating-" & channelid & "-" & infoid & "-" & projectid & "-" & J & ".html"">" & J &"</a>"
						End If
						n=n+1 : if n>=10 then exit for
					 Next
					
					 PageStr=PageStr & " <a class=""next"" href=""rating-" & channelid & "-" & infoid &"-" & projectid & "-" & PageNum & ".html"">末页</a>"
					 pageStr=PageStr & " <span>共" & totalPut & "条记录,分" & PageNum & "页</span></td></tr></table>"
				     PageStr = PageStr & "</td></tr></table></div>"
			         ShowPage = PageStr
End Function


End Class
%>

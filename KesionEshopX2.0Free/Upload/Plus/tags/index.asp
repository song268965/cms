<%Option Explicit%>
<!--#include File="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Template.asp"-->
<%

'******************************************************************
' Software name:KesionCMS X2.0
' Email: service@kesion.com . 营销QQ:4000080263  Tel:400-008-0263
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'******************************************************************
Dim KSCls
Set KSCls= New Tags
KSCls.Kesion()
Set KSCls = Nothing
Const MaxPerPage=10  '每页显示条数
Const MaxTags=500     '默认显示tags个数

Class Tags
    Private KS,KMR,F_C,LoopContent,SearchResult,photourl
	Private ChannelID,ClassID,SearchType,TagsName,SearchForm
    Private I,TotalPut, RS ,XML,Node,CurrPage,KeyTags
   
	Private Sub Class_Initialize()
		Set KS=New PublicCls
		Set KMR=New Refresh
	End Sub

	Private Sub Class_Terminate()
        closeconn
	    Set KS=Nothing
		Set KMR=Nothing
	End Sub
  
 Sub Kesion()
		   FCls.RefreshType = "tags" '设置刷新类型，以便取得当前位置导航等
		   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		   
			TagsName=KS.CheckXSS(KS.S("n"))
			If TagsName="" Then 
			 Call TagsMain()
			 F_C = KMR.KSLabelReplaceAll(F_C) 
			 Response.Write F_C
			Else
			 Call TagsList()
			End If
		   
 End Sub
 
 Sub TagsMain()
	F_C = KMR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & "common/tags.html")
	If Trim(F_C) = "" Then F_C = "模板不存在!"

   Dim TP:Tp=LFCls.GetConfigFromXML("tags","/labeltemplate/label","tags")
   Dim RS,SQL,K,str,Turl
   If InStr(tp,"{$ShowHotTags}")<>0 Then
	   Set RS=Conn.Execute("Select top " & MaxTags & " KeyText,hits From KS_KeyWords order by hits desc,id desc")
	   If Not RS.Eof Then SQL=RS.GetRows(-1)
	   RS.Close:Set RS=Nothing
	   If IsArray(SQL) Then
		 For k=0 to Ubound(SQL,2)
		  turl=KS.TagsUrl(SQL(0,K),0,0,1)
		  str=str & "<a href='" & Turl & "' title='已被使用了" & SQL(1,K) & "次'>" & SQL(0,K) & "</a>  "
		 Next
	   End If
	   Tp=Replace(Tp,"{$ShowHotTags}",str)
   End If
   
    If InStr(tp,"{$ShowNewTags}")<>0 Then
	   str=""
	   Set RS=Conn.Execute("Select top " & MaxTags & " KeyText,hits From KS_KeyWords order by adddate desc")
	   If Not RS.Eof Then SQL=RS.GetRows(-1)
	   RS.Close:Set RS=Nothing
	   If IsArray(SQL) Then
		 For k=0 to Ubound(SQL,2)
		    turl=KS.TagsUrl(SQL(0,K),0,0,1)
		  str=str & "<a href='" & Turl & "' title='已被使用了" & SQL(1,K) & "次'>" & SQL(0,K) & "</a>  "
		 Next
	   End If
	   Tp=Replace(Tp,"{$ShowNewTags}",str)
	End If
   
    F_C=Replace(F_C,"{$ShowTags}",Tp)
	F_C=Replace(F_C,"{$TagsName}","关键字Tags")
 End Sub
 

 Sub TagsList()
    SearchTags()
	F_C=Replace(F_C,"{$ShowTags}",SearchResult)
	F_C = Replace(F_C,"{$TagsName}",TagsName)
	F_C = Replace(F_C,"{$ShowTotal}",totalput)
  End Sub
  
  Sub TagsHits(ID)
    If ID<>0 Then
	 Conn.Execute("Update KS_KeyWords set hits=hits+1,lastusetime=" & SqlNowString & " where ID=" & ID)
	End IF
  End Sub
  
  Sub SearchTags() 

   
   Dim Param,TemplateID
   If IsNumeric(TagsName) Then
     Param=" Where ID=" & TagsName
   Else
     Param=" Where KeyText='" & TagsName &"'"
   End If
   
   CurrPage=KS.ChkClng(Request("Page"))
   ChannelID=KS.ChkClng(Request("ChannelID"))
   ClassID=KS.S("ClassID")


   Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "select top 1 * From KS_KeyWords" & Param,conn,1,1
   If Not RS.Eof Then
      TemplateID=RS("TemplateID")
	  KeyTags=RS("KeyText")
	  If CurrPage=1 Then Call TagsHits(RS("ID"))
   Else
      KeyTags=TagsName
   End If
   RS.Close
   
   Param=" Where DelTF=0 And Verific=1 And keywords like '%" & KeyTags & "%'"
   if ClassiD<>"" and ClassiD<>"0" then
		 Param=Param & " And Tid In(" & KS.GetFolderTid(ClassiD) & ")"
   end if
   
   'If KS.IsNul(TemplateID) and ClassID<>"" and classid<>"0" Then  TemplateID=KS.Setting(3) & KS.Setting(90) & "common/tagsList_" & split(KS.C_C(ClassID,8),",")(0) &".html"
   If KS.IsNul(TemplateID) Then TemplateID=KS.Setting(3) & KS.Setting(90) & "common/tagsList.html"
   
   F_C = KMR.LoadTemplate(TemplateID)
   
    Dim SqlStr
   If Channelid=0 Then
     SQLStr="select * From KS_ItemInfo " & Param
   Else
     SQLStr="select *," & ChannelID & " AS ChannelID From " & KS.C_S(ChannelID,2) &" " & Param
   End If
   
  'ks.echo sqlstr

  Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open SqlStr,Conn,1,1

  IF RS.Eof And RS.Bof Then
      totalput=0
      SearchResult = "Tags:<Font color=red>" & TagsName & "</font>,没有找到任何相关信息!"
  Else
					TotalPut= RS.Recordcount
                    If CurrPage > 1 and (CurrPage - 1) * MaxPerPage < totalPut Then
                            RS.Move (CurrPage - 1) * MaxPerPage
                    End If
					 Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","root")
    End IF
	RS.Close
	Set RS=Nothing
	
	F_C = KMR.KSLabelReplaceAll(F_C)
	 Scan F_C
  End Sub   

  
  
  Sub ParseArea(sTokenName, sTemplate)
			Select Case sTokenName
				Case "loop"
				      If IsObject(XML) Then
					   For Each Node In Xml.DocumentElement.SelectNodes("row")
						Scan sTemplate
					   Next
					  Else
					   echo "<div class='border' style='text-align:center'>对不起,根据您的查找条件,找不到任何相关记录!</div>"
					  End If
			End Select 
        End Sub 
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			    case "item" EchoItem sTokenName
				case "search" 
				          select case sTokenName
						    case "showpage" 
								If KS.Setting(185)="1" Then
								 echo ReplacePage(2,CurrPage,TotalPut,MaxPerPage)
								Else
								 echo KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
								End If
							case "totalput" echo TotalPut
							case "leavetime" 
							   dim leavetime:leavetime=FormatNumber((timer-starttime),5)
							   if leavetime<1 then leavetime="0"&leavetime
							   echo leavetime
							case "keyword" echo KS.R(KeyTags)
							case "channelid" echo channelid
							case "relatekeyword" relatekeyword
						  end select
			End Select
		End Sub
		Sub EchoItem(sTokenName)
		  Select Case sTokenName
		    case "id" echo GetNodeText("id")
			case "linkurl" 
			     if channelid=0 Then
				    echo KS.GetItemURL(GetNodeText("channelid"),GetNodeText("tid"),GetNodeText("infoid"),GetNodeText("fname"),GetNodeText("adddate"))
				 Else
				    echo KS.GetItemURL(GetNodeText("channelid"),GetNodeText("tid"),GetNodeText("id"),GetNodeText("fname"),GetNodeText("adddate"))
				 End If
			case "classname" 
			   echo KS.C_C(GetNodeText("tid"),1)
			case "classurl" 
			  echo KS.GetFolderPath(GetNodeText("tid"))
			case "intro" 
			 Dim Intro:intro=KS.Gottopic(KS.LoseHtml(GetNodeText("intro")),160)
			 Intro=Replace(Intro,"&nbsp;","")
			 If Not KS.IsNul(KeyTags) Then
			  echo Replace(Intro,KeyTags,"<span style='color:red'>" & KeyTags & "</span>")
			 Else
			 echo intro
			 End If
			case "keywordlist"
			  Call GetKeyWordList(GetNodeText("keywords"),GetNodeText("channelid"),GetNodeText("tid"))
			case else
			  echo GetNodeText(sTokenName)
		  End Select
		End Sub
		Function GetNodeText(NodeName)
		 Dim N,Str
		 NodeName=Lcase(NodeName)
		 If IsObject(Node) Then
		  set N=node.SelectSingleNode("@" & NodeName)
		  If Not N is Nothing Then Str=N.text
		  If Not KS.IsNul(KeyTags)  And NodeName="title" Then
			Str=Replace(Str,KeyTags,"<span style='color:red'>" &KeyTags & "</span>")
		  End If
		  GetNodeText=Str
		 End If
		End Function
  
         Sub GetKeyWordList(KeyWords,channelid,tid)
		   Dim cid:cid=KS.C_C(Tid,8)
		   if ks.isnul(cid) then exit sub
		   Dim TN:TN=Split(cid,",")(0)
		   Dim KeyArr:KeyArr=Split(KeyWords,",")
		   Dim I
		   For i=0 To Ubound(KeyArr)
		     Dim ID:ID=KS.Tags(KeyArr(i))
			 If IsNumeric(ID) Then
			   Echo "<a href=""" & KS.TagsUrl(ID,ChannelID,TN,1) &""" target=""_blank"">" & KeyArr(i) &"</a> "
			 Else
			   
			 End If
		   Next
		 End Sub
		 
		 
		'相关词条
		Sub RelateKeyword()
		  If IsObject(Application(KS.SiteSN&"_ClassTags")) Then
		    Dim Node,KK
		    For Each Node In Application(KS.SiteSN&"_ClassTags").DocumentElement.SelectNodes("row")
			  KK=node.selectsinglenode("@keytext").text
			  if instr(KK,KeyTags)<>0 Then
			  echo "<a href=""" & KS.TagsUrl(KK,0,0,1) &""" target=""_blank"">" & KK &"</a> "
			  End If
			Next
		  End If
		End Sub 
 
       '获取分页链接
		Function GetPageUrl(CurrPage)
			GetPageUrl=KS.TagsUrl(KS.S("N"), ChannelID, ClassID ,CurrPage)
		End Function
		
		Function ReplacePage(PageStyle,CurrPage,TotalPut,PerPageNumber)
		 Dim n:n=PageStyle
         Dim Tp,TotalPage,ItemUnit
		 Dim XML:Set XML=LFCls.GetXMLFromFile("pagestyle")
		 Dim Node:Set Node= XML.documentElement.selectSingleNode("/pagestyle/item[@name='" & n & "']/content")
		  If Not Node Is Nothing Then
		   Tp=Node.text
		  End If
		  
		  if (TotalPut mod PerPageNumber)=0 then
				TotalPage= TotalPut \ PerPageNumber
		  else
				TotalPage = TotalPut \ PerPageNumber + 1
		 end if
		 ItemUnit=KS.C_S(ChannelID,4): if KS.IsNul(ItemUnit) Then ItemUnit="条"
		  
		  
		  Dim homeUrl,endUrl,prevUrl,nextUrl 

            if (CurrPage = 1 and CurrPage <>TotalPage) then
                homeUrl = "javascript:;"
                prevUrl = "javascript:;"
                nextUrl = GetPageUrl(CurrPage + 1)
                endUrl = GetPageUrl(TotalPage)
            elseif (CurrPage = 1 and CurrPage = TotalPage) then
                homeUrl = "javascript:;"
                prevUrl = "javascript:;"
                nextUrl = "javascript:;"
                endUrl = "javascript:;"
            elseif (CurrPage = TotalPage and  CurrPage <> 2)  then '对于最后一页刚好是第二页的要做特殊处理
                homeUrl = GetPageUrl(1)
                prevUrl = GetPageUrl(CurrPage - 1)
                nextUrl = "javascript:;"
                endUrl = "javascript:;"
            elseif (CurrPage = TotalPage and CurrPage = 2) then
                homeUrl = GetPageUrl(1)
                prevUrl = GetPageUrl(1)
                nextUrl = "javascript:;"
                endUrl = "javascript:;"
            elseif (CurrPage = 2) then
                homeUrl = GetPageUrl(1)
                prevUrl = GetPageUrl(1)
                nextUrl = GetPageUrl(CurrPage + 1)
                endUrl = GetPageUrl(TotalPage)
            else
                homeUrl = GetPageUrl(1)
                prevUrl = GetPageUrl(CurrPage - 1)
                nextUrl = GetPageUrl(CurrPage + 1)
                endUrl = GetPageUrl(TotalPage)
            end if
		 
		  Tp=Replace(Tp,"{$homeurl}",homeurl)
		  Tp=Replace(Tp,"{$prevurl}",prevurl)
		  Tp=Replace(Tp,"{$nexturl}",nexturl)
		  Tp=Replace(Tp,"{$endurl}",endurl)

          if (instr(Tp,"{$pagenumlist}")<>0) then
                Dim j,p,pageStr:pageStr=""
                Dim StartPage:startpage = 1
                if (CurrPage >= 7)  then startpage = CurrPage - 5
                if (TotalPage - CurrPage < 5) then startpage = TotalPage - 9
                if (startpage <= 0) then startpage = 1
                Dim nn:nn = 1
                for p = startpage to TotalPage
                    if (p = CurrPage) then
                        pageStr=pageStr & " <a href=""#"" class=""curr"">" & p & "</a>"
                    else
                         pageStr=pageStr & " <a class=""num"" href=""" & GetPageUrl(p)& """>" & p & "</a>"
                    end if
					if (nn >= 10) then exit for
					nn=nn+1
                Next
                Tp = replace(Tp, "{$pagenumlist}", pagestr)
         End If
		 
		 if (instr(Tp,"{$turnpage}")<>0) then
                pageStr="<select name=""page"" id=""turnpage"" onchange=""javascript:window.location=this.options[this.selectedIndex].value;"">"
                for j = 1 to totalPage
                  pageStr=pageStr &"<option value=""" & GetPageUrl(j) & """"
				  if j=currPage then pageStr=pageStr &" selected"
				  pageStr=pageStr &">第" & j & "页</option>"
                next
                pageStr=pageStr &"</select>"
                Tp = replace(Tp, "{$turnpage}", pageStr)
        end if

		  
		 Tp=Replace(Tp,"{$currentpage}",CurrPage)
		 Tp=Replace(Tp,"{$maxperpage}",PerPageNumber)
		 Tp=Replace(Tp,"{$totalpage}",TotalPage)
		 Tp=Replace(Tp,"{$totalput}",TotalPut)
		 Tp=Replace(Tp,"{$itemunit}",ItemUnit)
		 ReplacePage=Tp
		End Function
  
  
  
  
  
End Class
%> 
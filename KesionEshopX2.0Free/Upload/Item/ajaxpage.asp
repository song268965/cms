<%@ Language="VBSCRIPT" codepage="65001" %>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KS:Set KS=New PublicCls
Dim KMRFObj:Set KMRFObj= New RefreshFunction

dim CurrPage,RS,SqlStr,iCurPage,ipresize,icount,ipagecount,ChannelID,str

CurrPage=KS.ChkClng(KS.S("curpage"))
If CurrPage<=0 Then CurrPage=CurrPage+1

Dim L_P,PageParamArr,LabelStyle,LabelParamStr
Dim LabelID:LabelID=KS.S("LabelID")   '标签ID
Dim ClassID:ClassID=KS.S("ClassID")   '栏目ID
Dim RefreshType:RefreshType=KS.S("RefreshType")   '调用类型 
Dim SpecialID:SpecialID=KS.ChkClng(KS.S("SpecialID"))

IF (KS.IsNul(Request.ServerVariables("HTTP_REFERER"))) Then KS.Die "error!"


 Dim RCls:Set RCls=New Refresh
 Call RCls.LoadLabelToCache()    '加载标签
L_P=Replace(rcls.LabelXML.documentElement.selectSingleNode("labellist[@labelid='" & LabelID & "']").text,LabelID,"ajax")
 Set RCls=Nothing


LabelStyle         = KS.GetTagLoop(L_P)
If RefreshType<>"ChannelSpecial" Then
    LabelParamStr      = Replace(Replace(L_P, "{Tag:GetPageList", ""),"}" & LabelStyle&"{/Tag}", "")
	ChannelID = KS.ChkClng(KS.C_C(ClassID,12))
	If RefreshType="Folder" And ChannelID=0 Then KS.Echo "标签参数加载出错!"  : Response.End()
Else
    LabelParamStr      = Replace(Replace(L_P, "{Tag:GetLastSpecialList", ""),"}" & LabelStyle&"{/Tag}", "")
End If
If LabelParamStr="" Or IsNull(LabelParamStr) Then Response.End()

				 Dim XMLDoc,XMLSql
				 Dim ParamNode,IncludeSubClass,ModelID,OrderStr,PrintType,PageStyle,PicStyle,ShowPicFlag,FieldStr,Param
				 Dim PerPageNumber,TotalPut,PageNum,TempStr,TableName,ItemUnit,ItemName
				 Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				 If XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
					 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
					 ModelID         = ParamNode.getAttribute("modelid") : If Not IsNumeric(ModelID) Then ModelID=1
					 IncludeSubClass = ParamNode.getAttribute("includesubclass") 
					 PrintType       = ParamNode.getAttribute("printtype") : If Not IsNumeric(PrintType) Then PrintType=1
					 PageStyle       = ParamNode.getAttribute("pagestyle") : If PageStyle="" Or IsNull(PageStyle) Then PageStyle=1
					 PicStyle        = ParamNode.getAttribute("picstyle")
					 OrderStr        = ParamNode.getAttribute("orderstr") : If OrderStr="" Or IsNull(OrderStr) Then OrderStr="ID Desc"
					 ShowPicFlag     = ParamNode.getAttribute("showpicflag") : If ShowPicFlag="" Or IsNull(ShowPicFlag) Then ShowPicFlag=false
					 PerPageNumber   = ParamNode.getAttribute("num") : If Not IsNumeric(PerPageNumber) Then PerPageNumber=10
					 
					 If RefreshType<>"ChannelSpecial" Then
						 Param = " Where I.Verific=1 And I.DelTF=0"
						 If RefreshType="Folder" Then
							 If CBool(IncludeSubClass) = True Then 
							 Param= Param & " And I.Tid In (" & KS.GetFolderTid(ClassID) & ")" 
							 Else 
							 Param= Param & " And I.Tid='" & ClassID & "'"
							 End If
						 Else
						   ChannelID=ModelID
						   Param= Param & KS.GetSpecialPara(ModelID,SpecialID)
						 End If
						 Set KMRFObj.ParamNode=ParamNode
						 Call KMRFObj.LoadField(ChannelID,PrintType,PicStyle,ShowPicFlag,FieldStr,TableName,Param)
	                     If Lcase(Left(Trim(OrderStr),2))<>"id" Then  OrderStr=OrderStr & ",I.ID Desc"
						 If ChannelID=0 Then
						   TableName="[KS_ItemInfo]" :   ItemUnit="条记录" : ItemName=""
						 Else
						   TableName=KS.C_S(ChannelID,2) : ItemUnit = KS.C_S(ChannelID,4) : ItemName=KS.C_S(ChannelID,3)
						 End If
				
						SqlStr = "SELECT " & FieldStr & " FROM " & TableName & " I " & Param & " ORDER BY I.IsTop Desc," & OrderStr
					Else
					  SqlStr="Select S.specialid,S.classid,S.SpecialName,S.SpecialEname,S.FsoSpecialIndex,S.SpecialAddDate as AddDate,S.PhotoUrl,S.SpecialNote As Intro,S.creater,C.ClassName as SpecialClassName From KS_Special S Inner Join KS_SpecialClass C On S.ClassID=C.ClassID Where S.ClassID=" & ClassID & " Order By S.SpecialID Desc"
					End If
					Set RS=Server.CreateObject("ADODB.RECORDSET")
					RS.Open SqlStr, Conn, 1, 1
					If RS.EOF And RS.BOF Then
						TempStr = "<p>此栏目下没有信息</p>"
					Else
						PerPageNumber=cint(PerPageNumber)
						If RefreshType<>"ChannelSpecial" Then
						 TotalPut = Conn.Execute("select Count(id) from " & TableName & " I " & Param)(0)
						Else
						 TotalPut = Conn.Execute("Select count(S.specialid) From KS_Special S Inner Join KS_SpecialClass C On S.ClassID=C.ClassID Where S.ClassID=" & ClassID)(0)
						End If
						if (TotalPut mod PerPageNumber)=0 then
								PageNum = TotalPut \ PerPageNumber
						else
								PageNum = TotalPut \ PerPageNumber + 1
						end if
						If CurrPage >1 and (CurrPage - 1) * PerPageNumber < totalPut Then
							RS.Move (CurrPage - 1) * PerPageNumber
						Else
							CurrPage = 1
						End If
						Set XMLSQL=KS.ArrayToXml(RS.GetRows(PerPageNumber),RS,"row","root")
						Call KMRFObj.LoadPageParam(XMLSQL,ParamNode,ChannelID)
						If RefreshType<>"ChannelSpecial" Then
						 str=KMRFObj.ExplainGerericListLabelBody(LabelStyle)
						Else
						 str=KMRFObj.ExplainSpecialListLabelBody(LabelStyle)
						End If
						XMLSql=Empty
						
					End If
				
					RS.Close:Set RS=Nothing					
					XMLDoc= Empty : Set ParamNode=Nothing
                End iF


Dim Tp,homeUrl,endUrl,prevUrl,nextUrl,startpage
Dim XML:Set XML=LFCls.GetXMLFromFile("pagestyle")
Dim Node:Set Node= XML.documentElement.selectSingleNode("/pagestyle/item[@name='" &PageStyle & "']/content")
If Not Node Is Nothing Then
		 Tp=Node.text
End If
 tp=replace(tp,"{$maxperpage}",PerPageNumber)
 tp=replace(tp,"{$totalput}",TotalPut)
 tp=replace(tp,"{$totalpage}",PageNum)
 tp=replace(tp,"{$currentpage}",CurrPage)
 tp=replace(tp,"{$itemunit}",ItemUnit)
 
homeUrl = "javascript:homePage(" & CurrPage & ");"
if (currPage=1) then
 prevUrl = "javascript:homePage();"
else
 prevUrl = "javascript:turnPage(" & currpage-1 & ");"
end if
if (currpage=pagenum) then
 NextUrl = "javascript:lastPage(" & currpage & "," & PageNum &");"
 endUrl = "javascript:lastPage(" & currpage & "," & PageNum &");"
else
 NextUrl = "javascript:turnPage(" & currpage+1 & ");"
 endUrl="javascript:turnPage(" & pagenum & ");"
end if
				
				
 Tp=Replace(Tp,"{$homeurl}",homeurl)
 Tp=Replace(Tp,"{$prevurl}",prevurl)
 Tp=Replace(Tp,"{$nexturl}",nexturl)
 Tp=Replace(Tp,"{$endurl}",endurl) 

 if (instr(Tp,"{$pagenumlist}")<>0) then
                Dim j,p,pageStr:pageStr=""
                startpage = 1
                if (CurrPage >= 7)  then startpage = CurrPage - 5
                if (PageNum - CurrPage < 5) then startpage = pageNum - 10
                if (startpage <= 0) then startpage = 1
                Dim nn:nn = 1
                for p = startpage to pageNum
                    if (p = CurrPage) then
                        pageStr=pageStr & " <a href=""#"" class=""curr"">" & p & "</a>"
                    else
                         pageStr=pageStr & " <a class=""num"" href=""javascript:turnPage(" & p& ");"">" & p & "</a>"
                    end if
                    nn=nn+1
					if (nn >= 10) then exit for
                Next
                Tp = replace(Tp, "{$pagenumlist}", pagestr)
End If	  

if (instr(Tp,"{$turnpage}")<>0) then
                pageStr="<select name=""page"" id=""turnpage"" onchange=""javascript:turnPage(this.options[this.selectedIndex].value);"">"
                for j = 1 to pageNum
                  pageStr=pageStr &"<option value=""" & j & """"
				  if j=currPage then pageStr=pageStr &" selected"
				  pageStr=pageStr &">第" & j & "页</option>"
                next
                pageStr=pageStr &"</select>"
                Tp = replace(Tp, "{$turnpage}", pageStr)
    end if
		

 if instr(str,"[KS:PageStyle]")=0 then
  str=str  &"[KS:PageStyle]"
 end if
 

response.write replace(str,"[KS:PageStyle]",tp)



set KS=Nothing
CloseConn
%>

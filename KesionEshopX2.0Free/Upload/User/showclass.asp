<!--#include file="../conn.asp"-->
<!--#include file="../ks_cls/kesion.membercls.asp"-->
<%
const tj=1    '从第几级算起
Dim KS:Set KS=new PublicCls
Dim KSUser:Set KSUser=New UserCls
'If KSUser.UserLoginChecked=false Then   KS.Die ""
Dim SQL,K,Node,Pstr,Xml,ChannelID
ChannelID=KS.ChkClng(KS.S("ChannelID"))
KS.LoadClassConfig()
dim n:n=0

dim classid:classid=ks.s("classid")
dim FieldName:FieldName=KS.S("FieldName")
If KS.IsNul(FieldName) Then FieldName="ClassID"
if classid="" then classid="0"
%>
var subsmallclassid<%=FieldName%> = new Array();
<%
If ChannelID<>0 Then Pstr="and @ks12=" & channelid & ""
 Set Xml=Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
 For Each Node In Xml
        If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,KSUser.GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3)) Or Node.SelectSingleNode("@ks20").text="0" Then
		%>
		subsmallclassid<%=FieldName%>[<%=n%>] = new Array('<%=Node.SelectSingleNode("@ks13").text%>','<%=Node.SelectSingleNode("@ks0").text%>','<%=Node.SelectSingleNode("@ks1").text%>',0,<%=Node.SelectSingleNode("@ks19").text%>)
		<%
	    Else
		%>
		subsmallclassid<%=FieldName%>[<%=n%>] = new Array('<%=Node.SelectSingleNode("@ks13").text%>','<%=Node.SelectSingleNode("@ks0").text%>','<%=Node.SelectSingleNode("@ks1").text%>',1,<%=Node.SelectSingleNode("@ks19").text%>)
		<%
		End IF
         n=n+1
 Next

%>
function changesmallclassid<%=FieldName%>(selectValue)
{
if (selectValue==0) return;
document.getElementById('smallerclassid<%=FieldName%>').length = 0;   //点击一级栏目时,置三级下拉为空
document.getElementById('smallerclassid<%=FieldName%>').options[0] = new Option('请选择...','0');

document.getElementById('smallclassid<%=FieldName%>').length = 0;
document.getElementById('smallclassid<%=FieldName%>').options[0] = new Option('请选择...','0');

document.getElementById('<%=FieldName%>').value='0';

	  document.getElementById('smallclassid<%=FieldName%>').style.display='';
	  document.getElementById('smallerclassid<%=FieldName%>').style.display='';


for (i=0; i<subsmallclassid<%=FieldName%>.length; i++)
{
    if (subsmallclassid<%=FieldName%>[i][1] == selectValue && subsmallclassid<%=FieldName%>[i][4]==0){  //只有一级的情况
	  document.getElementById('<%=FieldName%>').value=selectValue; 
	  document.getElementById('smallclassid<%=FieldName%>').style.display='none';
	  document.getElementById('smallerclassid<%=FieldName%>').style.display='none';
	  return;
	}else if (subsmallclassid<%=FieldName%>[i][0] == selectValue)
	{
	     //判断有没有下级允许投稿
		 var xjtk=false;
		 for(j=0;j< subsmallclassid<%=FieldName%>.length; j++)
		 {
		    if (subsmallclassid<%=FieldName%>[j][0]==subsmallclassid<%=FieldName%>[i][1]){
			  if (subsmallclassid<%=FieldName%>[j][3]==1){
			    xjtk=true;
				break;
			  }
			}
		 }
	     if (subsmallclassid<%=FieldName%>[i][3] == 1 || xjtk ){
			document.getElementById('smallclassid<%=FieldName%>').options[document.getElementById('smallclassid<%=FieldName%>').length] = new Option(subsmallclassid<%=FieldName%>[i][2], subsmallclassid<%=FieldName%>[i][1]);
		 }
		 
		 //判断是否显示三级下拉列表
		 var showxj=false;
		 for(j=0;j< subsmallclassid<%=FieldName%>.length; j++){
		    if (subsmallclassid<%=FieldName%>[j][0]==selectValue){
			   if (parseInt(subsmallclassid<%=FieldName%>[j][4])>0){
			    showxj=true;
				break;
			   }
			}
		 }
		 if (showxj==true){
		 document.getElementById('smallerclassid<%=FieldName%>').style.display='';
		 }else{
		 document.getElementById('smallerclassid<%=FieldName%>').style.display='none';
		 }
		 
	}
}
}
function changesmallerclassid<%=FieldName%>(selectValue)
{
if (selectValue=='0') document.getElementById('<%=FieldName%>').value='0';
//判断是否显示三级下拉列表
for (i=0; i<subsmallclassid<%=FieldName%>.length; i++){
     if (subsmallclassid<%=FieldName%>[i][1]==selectValue){
	  	  if (subsmallclassid<%=FieldName%>[i][4]==0){
		     document.getElementById('<%=FieldName%>').value=selectValue;
		     document.getElementById('smallerclassid<%=FieldName%>').style.display='none';
		  }else{
		     document.getElementById('<%=FieldName%>').value='0';
			 document.getElementById('smallerclassid<%=FieldName%>').style.display='';
		}
	 }
}

document.getElementById('smallerclassid<%=FieldName%>').length = 0;
document.getElementById('smallerclassid<%=FieldName%>').options[0] = new Option('请选择...','0');
for (i=0; i<subsmallclassid<%=FieldName%>.length; i++)
{

	if (subsmallclassid<%=FieldName%>[i][0] == selectValue)
	{

		if (subsmallclassid<%=FieldName%>[i][3] == 1){
		document.getElementById('smallerclassid<%=FieldName%>').options[document.getElementById('smallerclassid<%=FieldName%>').length] = new Option(subsmallclassid<%=FieldName%>[i][2], subsmallclassid<%=FieldName%>[i][1]);
		}
		
	}
}
}
function getclassid(selectValue){
document.getElementById('<%=FieldName%>').value=selectValue;
<%if channelid=5 then response.write "getBrandList();"%>
}

document.write ("<select class='select' name='bigclassid' modified='false' id='bigclassid' style='width:120px' size='1' onChange='changesmallclassid<%=FieldName%>(this.value)'>");
document.write ("<option value='0' selected>请选择...</option>");
<%
 Set Xml=Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&" and @ks10=" & tj & "]")
 For Each Node In Xml
  If ((Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,KSUser.GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3))) and checkxjtk(Node.SelectSingleNode("@ks0").text)=false Then
  Else%>
document.write ("<option value=<%=Node.SelectSingleNode("@ks0").text%>><%=Node.SelectSingleNode("@ks1").text%></option>");
<%
  End If
 Next
%>
document.write ("</select>")

document.write ("  <select class='select' modified='false' name='smallclassid<%=FieldName%>' size='1' onChange='changesmallerclassid<%=FieldName%>(this.value)' style='width:120px;display:none' id='smallclassid<%=FieldName%>'>");
document.write ("<option value='0' selected>请选择...</option>");
document.write ("</select>")
document.write ("  <select class='select' modified='false' name='smallerclassid<%=FieldName%>' size='1' style='display:none;width:120px' id='smallerclassid<%=FieldName%>' onChange='getclassid(this.value)'>");
document.write ("<option value='0' selected>请选择...</option>");
document.write ("</select>");
document.write ("<input type='hidden' name='<%=FieldName%>' value='<%=classid%>' id='<%=FieldName%>'/>");
<%

'默认值
If ClassID<>"0" Then
 If KS.C_C(ClassID,10)-tj=0 Then   '一级
 %>
 	document.getElementById('bigclassid').value='<%=ClassID%>';
    document.getElementById('smallclassid<%=FieldName%>').style.display='none';
    document.getElementById('smallerclassid<%=FieldName%>').style.display='none';
 <%
 ElseIf KS.C_C(ClassID,10)-tj=1 Then   '二级
 %>
    document.getElementById('smallclassid<%=FieldName%>').style.display='';
	setSecoundOption('<%=KS.C_C(ClassID,13)%>','<%=ClassID%>');
	document.getElementById('bigclassid').value='<%=KS.C_C(ClassID,13)%>';
 <%
 ElseIf KS.C_C(ClassID,10)-tj=2 Then   '三级
   %>
    document.getElementById('smallerclassid<%=FieldName%>').style.display='';
	for (i=0; i<subsmallclassid<%=FieldName%>.length; i++){
	   //给三级下拉指定值
      if (subsmallclassid<%=FieldName%>[i][0]=='<%=KS.C_C(ClassID,13)%>'){
		if (subsmallclassid<%=FieldName%>[i][3] == 1){
		document.getElementById('smallerclassid<%=FieldName%>').options[document.getElementById('smallerclassid<%=FieldName%>').length] = new Option(subsmallclassid<%=FieldName%>[i][2], subsmallclassid<%=FieldName%>[i][1]);
		}
	   }
	   //得二级下拉的ParentID
	   if (subsmallclassid<%=FieldName%>[i][1]=='<%=KS.C_C(ClassID,13)%>'){
	    pid=subsmallclassid<%=FieldName%>[i][0];
	   }
    }
	document.getElementById('smallerclassid<%=FieldName%>').value='<%=ClassID%>';
	
	//给二级下拉指定值
	setSecoundOption(pid,'<%=KS.C_C(ClassID,13)%>');
	document.getElementById('bigclassid').value=pid;
   <%
 End If
 %>
<%End If
Set KS=Nothing
Set KSUser=Nothing
CloseConn
%>

//给二级下拉填充值 参数pid 父栏目ID, sid 选中的栏目ID
function setSecoundOption(pid,sid)
{
	//给二级下拉指定值
	for (i=0; i<subsmallclassid<%=FieldName%>.length; i++){
	if (subsmallclassid<%=FieldName%>[i][0] == pid)
	{
	     //判断有没有下级允许投稿
		 var xjtk=false;
		 for(j=0;j< subsmallclassid<%=FieldName%>.length; j++)
		 {
		    if (subsmallclassid<%=FieldName%>[j][0]==subsmallclassid<%=FieldName%>[i][1]){
			  if (subsmallclassid<%=FieldName%>[j][3]==1){
			    xjtk=true;
				break;
			  }
			}
		 }
	     if (subsmallclassid<%=FieldName%>[i][3] == 1 || xjtk ){
			document.getElementById('smallclassid<%=FieldName%>').options[document.getElementById('smallclassid<%=FieldName%>').length] = new Option(subsmallclassid<%=FieldName%>[i][2], subsmallclassid<%=FieldName%>[i][1]);
		 }
		 
	 }
	}
	document.getElementById('smallclassid<%=FieldName%>').value=sid;

}
<%
'检查栏目ID检查下级有没有允许投稿的栏目
function checkxjtk(id)
     Dim Xml,Node
	 Set Xml=Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&" and @ks10>" & tj & "]")
	 For Each Node In Xml
	   If KS.FoundInArr(Node.SelectSingleNode("@ks8").text,id,",")=true Then  '如果是他的下级
		  If ((Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,KSUser.GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3))) Or Node.SelectSingleNode("@ks20").text="0" Then
		  Else
		   checkxjtk=true
		   exit function
		  End If
	   End If
	Next

  checkxjtk=false
end function
%>
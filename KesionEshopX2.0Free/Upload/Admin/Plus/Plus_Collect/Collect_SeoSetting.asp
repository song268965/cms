<!--#include file="../../../conn.asp"-->
<!--#include file="../../../ks_cls/kesion.commoncls.asp"-->
<%
   Dim config,i,fs,f 

dim ks:set ks=new publiccls
if ks.s("action")="save" then
	for i=0 to 10
	   config=config&KS.G("config"&i)&"^`^"
	next
   call ks.WriteTOFile(KS.Setting(3)& KS.Setting(89) & "plus/plus_collect/seo/config.txt", config)
   call ks.WriteTOFile(KS.Setting(3)& KS.Setting(89) & "plus/plus_collect/seo/tyc.txt", request.Form("tyc"))
   ks.alerthintscript("恭喜,成功保存!")
end if
dim tyc
tyc=KS.ReadFromFile(KS.Setting(3)& KS.Setting(89) & "plus/plus_collect/seo/tyc.txt")

config=split(KS.ReadFromFile(KS.Setting(3)& KS.Setting(89) & "plus/plus_collect/seo/config.txt"),"^`^")




%>
<!DOCTYPE HTML><html><head><title>采集系统</title><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><link rel="stylesheet" type="text/css" href="../../Include/Admin_Style.css"><script src="../../../ks_inc/jquery.js"></script></head><body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class="topdashed" style="text-align:center"><strong>SEO 辅助工具设置</strong></div>
<div class="pageCont2">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable" ><form method="post" action="?action=save" name="myform"  onSubmit="return(CheckForm(this))">    
  <tr class='tdbg'>
    <td height="25" align="right" class='clefttitle'> <strong>同义词替换:</strong>
	<br>
	<font color=green>说明：启用此功能会影响采集速度，同义词库越大，速度越慢。</font>
	</td>
    <td height="25" class='clefttitle'  style="text-align:left">
	 <table border="0">
	  <tr>
	   <td nowrap style="text-align:left">
	 标题替换:
	<label>
        <input type="radio" name="config0" value="1"  <% if config(0)=1 Then %>checked="checked"<% End if%>/>
        开启
        <input type="radio" name="config0" value="0"  <% if config(0)=0 Then %>checked="checked"<% End if%>/>
        关闭
      </label>
	  <br/>
	 正文替换:
	 <label>
        <input type="radio" name="config1" value="1"  <% if config(1)=1 Then %>checked="checked"<% End if%>/>
        开启
        <input type="radio" name="config1" value="0"  <% if config(1)=0 Then %>checked="checked"<% End if%>/>
        关闭
      </label>
	  <br/>
	  </td>
	  <td>
	  	 <font color=blue>开启同义词替换后,系统将采集到的数据根据同义词库进行替换。</font>
	  </td>
	  </tr>
	  <tr>
	   <td colspan=2>
	    	  <label><input type="checkBox" <%if config(2)="1" then response.write " checked"%> value="1" name="config2">采用双向替换</label>
			  <br/>
            启用双向替换速度会相对较慢,效果如下：<br/>
			如同义词：公司=企业  <br/>
			<strong>原文：</strong>科汛公司是国内较早的CMS开发企业。<br/>
			<strong>单向：</strong>科汛<font color=red>企业</font>是国内较早的CMS开发企业。<br/>
			<strong>双向：</strong>科汛<font color=red>企业</font>是国内较早的CMS开发<font color=red>公司</font>。<br/>
			
	   <td>
	  </tr>
	 </table>
	</td>
  </tr>
  <tr class='tdbg'>     <td height="25" colspan="2" class='clefttitle'  style="text-align:left">说明:一行一个规则,格式如下:<font color=red><br/>被替换的词1=同义词2<br/>被替换的词2=同义词2</font></td>     
</tr>     <tr class='tdbg'>    <td  width="20%" height="25" align="center" class='clefttitle'><strong>同义词替换规则：</strong>
<br/><font color=green>词库在"管理目录plus\KS_Collect\Seo\tyc.txt"，请根据自已的需要修改</font>
</td>     <td width="796"><textarea name="tyc" cols="49" rows="30"><%=tyc%></textarea></td>   </tr>
  <tr class='tdbg'>
    <td colspan=2 style='text-align:center'>
	 <input type='submit' value='确定保存' class='button'>
	</td>
  </tr>
  </form></table>
  </div>
</body></html>
/*================================================================
Created:2015-10-27
Copyright:www.Kesion.com  bbs.kesion.com
Version:KesionCMS X1.5
Service QQ：4000080263
==================================================================*/
function Page(curPage,labelid,classid,installdir,url,refreshtype,specialid,infoid)
{
   this.labelid=labelid;
   this.classid=classid;
   this.url=url;
   this.c_obj="c_"+labelid;
   this.p_obj="p_"+labelid;
   this.installdir=installdir;
   this.refreshtype=refreshtype;
   this.specialid=specialid;
   this.infoid=infoid;
   this.page=curPage;
   loadFunData(1);
   }
function loadFunData(p){  
   this.page=p;
   $("#"+c_obj).html("<div align='center'>正在读取数据...</div>");
   $.ajax({
		  type:"post",
		  url:installdir+url+"?labelid="+escape(labelid)+"&infoid="+infoid+"&classid="+classid+"&refreshtype="+refreshtype+"&specialid=" +specialid+"&curpage="+p+getUrlParam(),
		  cache:false,
		  success:function(d){
			$("#"+c_obj).html("<ul>"+d+"</ul>");
  }});
}
function homePage(i)
{
   if(i==1)
    alert("已经是首页了！")
   else
   loadFunData(1);
} 
function lastPage(i,e)
{
   if(i==e)
    alert("已经是最后一页了！")
   else
   loadFunData(e);
} 
function turnPage(i)
{
     loadFunData(i);
}
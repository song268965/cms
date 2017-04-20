/*================================================================
Modify Date:2014-7-13
Author：linwenzhong
Copyright:www.Kesion.com  bbs.kesion.com
Version:KesionCMS X1.5
营销QQ：4000080263
==================================================================*/
//评论支持
function Support(channelid,infoid,postid,id,typeid,installdir) { 
   jQuery.getScript(installdir+'plus/digmood/Comment.asp?action=Support&channelid='+channelid+'&infoid='+infoid+'&postid='+postid+'&Type='+typeid+'&id=' +id,   
       function(){ 
	       if (json.message=='good'){
			   $("#d"+id).html(parseInt($("#d"+id).html())+1);
		   }else if (json.message=='bad'){
			   $("#c"+id).html(parseInt($("#c"+id).html())+1);
		   }else{
			   tipsShow(json.message);
		   }
        });     
}

 //回复
var box='';
function replyCmt(postId,channelid,infoid,quoteId,installdir){
	box=$.dialog.open(installdir+"plus/digmood/comment.asp?action=ShowQuote&channelid="+channelid+"&infoid="+infoid+"&quoteid="+quoteId+"&postId="+postId,{id:'quotebox',lock:true,title:'引用回复',width:500,height:240,min:false,max:false});
}
 //当前页,频道ID,栏目ID，信息ID,Action,InstallDir
function Page(curPage,channelid,infoid,action,maxperpage,installdir){
   this._channelid = channelid;
   this._infoid    = infoid;
   this._action    = action;
   this._maxperpage= maxperpage;
   this._url       = installdir +"plus/digmood/comment.asp";
   
   this._c_obj="c_"+infoid;
   this._p_obj="p_"+infoid;
   this._installdir=installdir;
   this._page=curPage;
     loadDate(1,0);
   }
function loadDate(p,postload){
    this._page=p;
    var loadurl=_url+"?postload="+postload+"&channelid="+_channelid+"&infoid="+_infoid+"&from3g="+from3g+"&maxperpage="+_maxperpage+"&action=" +_action+"&page="+p;
    jQuery.getScript(loadurl+'&printout=js',   
       function(){ 
	      show(json.message);
    });     
}
function show(s)
{ 	
  if (s.indexOf("ks:page")==-1){
	 $(".cmtnum").html(parseInt($(".cmtnum").html())+1);
	 $("#cc_"+this._infoid).prepend(s);
  }else{
	  var pagearr=s.split("{ks:page}")
	  var pageparamarr=pagearr[1].split("|");
	  count=pageparamarr[0];    
	  perpagenum=pageparamarr[1];
	  pagecount=pageparamarr[2];
	  itemunit=pageparamarr[3];   
	  itemname=pageparamarr[4];
	  pagestyle=pageparamarr[5];
	  pagestyle=1;
	  if (this._page>1){
	   $("#cc_"+this._infoid).append(pagearr[0]);
	  }else{
	   $("#"+_c_obj).html(pagearr[0]);
	  }
      pagelist();
  }
}

function pagelist(){
     var statushtml="";
	 if (parseInt(this.pagecount)<=parseInt(this._page)){
	 statushtml="<div id=\"cmtloadtips\" style=\"cursor:pointer;height:30px;line-height:30px;background:#888;font-weight:bold;color:#fff;text-align:center;\">已加载完全部内容</div>"
	 }
	 else{
	 statushtml="<div id=\"cmtloadtips\" style=\"cursor:pointer;height:30px;line-height:30px;background:#888;font-weight:bold;color:#fff;text-align:center;\" onclick=\"nextPage()\">加载更多内容</div>"
	 }

  if (this.pagecount!=""&&this.count!=0)
	 {
	 $("#"+this._p_obj).html('<div style="margin-top:8px">'+statushtml+'</div>');
	 }
}

function nextPage()
{
   if(this._page<this.pagecount){
      loadDate(this._page+1,0);
	  $("#cmtloadtips").html("正在加载中..");
   }
   else
      tipsShow("已经到最后一页了");
}

function tipsShow(str){
	 if (from3g==1){
		  alert(str);
	 }else{
		 KesionJS.Alert(str);
	 }
}
 function getCookie(objName){//获取指定名称的cookie的值
   var arrStr = document.cookie.split("; ");
   for(var i = 0;i < arrStr.length;i ++){
    var temp = arrStr[i].split("=");
    if(temp[0] == objName) return unescape(temp[1]);
   } 
  }

function tag(id){
   if ($("#t"+id).attr("class")=="ta"){
     $("#t"+id).attr("class","ta tao")
	 addCookie("sj"+id, "ta tao",0);
   }else{
     $("#t"+id).attr("class","ta")
	 addCookie("sj"+id, "ta",0);
   }
  }
  
  function initialCss(){
	   //单选，多选
		$(".dx_button").find("input").click(function(){
			$(this).blur();			
			$(this).parent().parent().find("label").attr("class","dx_button");
			if ($(this).attr("type")=="radio"){
		   		$(this).parent().attr("class","dx_button dx_button_bg");
			}else{ //多选
				$(this).parent().parent().find("label").find("input").each(function(){
						if ($(this).prop('checked')){
							 $(this).parent().attr("class","dx_button dx_button_bg");
						}else{
							 $(this).parent().attr("class","dx_button");
						}
				});
			}
		})
		$(".SJlist").hover(function (){$(this).attr('class','SJlist SJlist_b')},function (){$(this).attr('class','SJlist');});
 }
 
  function ExamSubmit(){
    getAnswer();
	$.dialog.confirm('确定交卷吗？',function(){
		 document.myform.timeout.value=timeout;
	  	 $("#myform").submit();
	  },function(){});
	
	
  }
  function ExamSubmitByTips(){
	  getAnswer();
	  var hasfinish=parseInt(getFinishTMS());
	  var tips='';
	  if (hasfinish<totalTms){
		  tips="本卷共有<span style='color:blue'>"+totalTms+"</span>道题，您还有<span style='color:red'>"+(totalTms-hasfinish)+"</span>道题未答，";
	  }
	$.dialog.confirm(tips+'确定交卷吗？',function(){
		 document.myform.timeout.value=timeout;
	  	 $("#myform").submit();
	  },function(){}); 
  }
  
  //返回已做题目数
  function getFinishTMS(){
	var ans='',temp='',tmtype='';
	var arr=null;
	var hasnum=0;
    jQuery(".hidans").each(function(){
	   temp=jQuery(this).val();
	   tmtype=jQuery(this).attr("tmtype");
		arr=temp.split('$,$');
	   
	   for(var i=0;i<arr.length;i++){
			   if (arr[i].replace(/※/g,'')!=''){
				    hasnum++;
			   }
		 }
       
    });
	return hasnum;
  }
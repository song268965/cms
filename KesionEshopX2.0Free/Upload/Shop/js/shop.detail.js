 <!--
 var box='';
 var shop={
		     check:function(a){
			  for(var k=0;k<=a;k++)
			  {
			   if ($("#attr"+k).val()==''){
			     alert('请选择'+$("#attrname"+k).html().replace('：','！'));
				 return false;
			   }
			  }
			  if ($("#num").val()==0||$("#num").val()==''){
			   alert('请输入购买数量!');
			   return false;
			  }
			  if (parseInt($("#num").val())>parseInt($("#stock").text())){
			   alert('对不起，您最多只能购买'+$("#stock").text()+'!');
			   return false;
			  }
			  return true;
			  
			 },
			 gobuyIsScore:function(a,s){
			  if (shop.check(a)){
				  $("#isshop").val(s);
				  $("#cartform").submit();
				  }
		     },
             gobuy:function(a){
			  if (shop.check(a)){$("#cartform").submit();}
		     },
			 checkchangebuy:function(changebuyneedprice){
				 var totalprice=parseFloat($("#totalprice").text());
				 if (totalprice<parseFloat(changebuyneedprice)){
					  alert('对不起,您的订单费用不足,还差￥' +(parseFloat(changebuyneedprice)-totalprice).toFixed(2)+' 元!');
					  return false;
				 }
				 return true;
			 },
			 buynum:function(f){
			   if (f==0&&parseInt($("#num").val())<=1){
				   alert('对不起，您最少要购买1'+$("#unit").text()+'!');
				    return;
			   }
			   if (f==0 && parseInt($("#num").val())>0){
				   $("#num").val(parseInt($("#num").val())-1);
				   return;
			   }
			   if (f==1 && (parseInt($("#num").val())+1)>parseInt($("#stock").text()))
			   {
			    alert('对不起，您最多只能购买'+$("#stock").text()+$("#unit").text()+'!');
			    return;
			   }else{
				   $("#num").val(parseInt($("#num").val())+1) ;
			   }
			 },
			 buynums:function(type,proid,attrid,cartid,buynum,f){
				if (buynum!=0){$("#Q_"+cartid)[0].disabled=true;}
				$.get(dir+"shop/ajax.getdate.asp",{action:'getcartstock',type:type,buynum:buynum,cartid:cartid,f:f,proid:proid,attrid:attrid},function(s){
				  s=unescape(s);
				  if (buynum!=0){$("#Q_"+cartid)[0].disabled=false;}
				   if (s!='succ'){
					  alert(s.split('|')[1]);
					  return;
				   }else{
					  if (f==-1){
					  }else if (f==0){
						 $("#Q_"+cartid).val(parseInt($("#Q_"+cartid).val())-1);
					  }else{
						 $("#Q_"+cartid).val(parseInt($("#Q_"+cartid).val())+1);
					  }
					  shop.changecart(cartid);
				   }
				});

			 },
			 changecart:function(cartid){
				var price=$("#hidmyprice"+cartid).text();
				var price_Member=$("#hidmyprice_Member"+cartid).text();
				var score=$("#hidmyscore"+cartid).text();
				var isscore=$("#hidisscore"+cartid).text();
				 var num=parseInt($("#Q_"+cartid).val());
				var wholesaleNum=parseInt($("#hidWholesaleNum"+cartid).val());
				var wholesalePrice=parseFloat($("#hidWholesalePrice"+cartid).val());
				 if (num>=wholesaleNum && wholesaleNum>0) {
					 price=wholesalePrice;
				 }
				 $("#realPrice"+cartid).html(price);
				 $("#realscore"+cartid).html(isscore*num);
				 $("#realmember"+cartid).html((price_Member*num).toFixed(2));
				
					  $("#myprice"+cartid).html((price*num).toFixed(2));
					  $("#myscore"+cartid).html(parseInt(score*num));//bug改动
					  var totalmyprice=0;
					  var totalmyscore=0;
					  $("#shoppingtable").find("SPAN[name=totalmyprice]").each(function(){
							totalmyprice+=parseFloat($(this).text());															
					  });
					  $("#shoppingtable").find("SPAN[name=totalmyscore]").each(function(){
							totalmyscore+=parseInt($(this).text());															
					  });
					  $("#totalprice").html(totalmyprice.toFixed(2));
					 $("#totalscore").html(parseInt(totalmyscore));//bug改动
			 },
			 getAttr:function(obj,i,a,l,num,ids){
				 if (obj.innerHTML.indexOf('IMG')==-1){
					$("#attr"+i).val(obj.innerHTML.replace(/<I><\/I>/g,"")); 
				 }else{
					$("#current_img").attr("src",$("#"+obj.id).find("IMG").attr("src"));
					$("#proimg").attr("href",$("#"+obj.id).find("IMG").attr("src"));
					$("#attr"+i).val($("#"+obj.id).find("IMG").attr("title"));
				 }
			  if (i==1){
			   $("#attr2").val('');
			   $("#attr3").val('');
			  }else if(a==3&&i==2){
			   $("#attr3").val('');
			  }
			  
			  for(var k=0;k<=l;k++)
			  {
			   $("#att"+i+k)[0].className='txt';
			  }
			  obj.className='curr';
			  
			  if (i==a){
               $("#attrid").val(ids);
			   $('.vipprice').html('￥' +itemattr[ids][3]+ '元');
			   $.getScript(dir+'shop/ajax.getdate.asp?action=getstock&attrid='+ids,function(){
				    $("#stock").html(data.amount);
					if (parseInt($("#num").val())>parseInt($("#stock").text())){
					 $("#num").val($("#stock").text());
				    }
				});

			  }else{
               $("#attrid").val('');
			  }
			  
			  if (num!=0&&ids!=undefined){
				var str2='';
				var str22='';
				var idsarr=ids.split(",");
				var maxprice=minprice=0;
				for(var k=0;k<idsarr.length;k++){
					 var id=idsarr[k];
					 var itemvalue=itemattr[id][1];
					 if (num==3){
						itemvalue=itemattr[id][2]; 
					 }
					  str2+=itemvalue+"^"+id+",,,";
					  if (str22.indexOf(itemvalue)==-1){
					    str22+=itemvalue+",,,";
					  }
					if (parseFloat(itemattr[id][3])<minprice||minprice==0){minprice=itemattr[id][3];}
					if (parseFloat(itemattr[id][3])>maxprice||maxprice==0){maxprice=itemattr[id][3];}
				}
				if (minprice==maxprice){
				   	$('.vipprice').html('￥' +minprice+ '元');
				}else{
					$('.vipprice').html('￥'+ minprice+'元~￥' +maxprice+'元');
				}
					
				str22arr=str22.split(',,,');
				str2arr=str2.split(',,,');
				var new2str='';
				for(var i=0;i<str22arr.length-1;i++){
				      var ids2="";
					  for(var k=0;k<str2arr.length-1;k++){
					    if (str22arr[i]==str2arr[k].split('^')[0]){
						  if (ids2==""){
						   ids2=str2arr[k].split('^')[1];
						  }else{
						   ids2=ids2 + "," +str2arr[k].split('^')[1]
						  }
						}
					  }
					  if (i==0){
					   new2str=str22arr[i] + "^" + ids2;
					  }else{
					   new2str=new2str + ",,," + str22arr[i] + "^" + ids2;
					  }	
				}
				
				new2str=new2str.split(',,,');
				var strs="<span id='attrname"+num+"'>"+ myitemname[num-1]+ "：</span>";
				var vlan=new2str.length-1;
				for(var k=0;k<=vlan;k++){
					
					itemvalue=new2str[k].split("^")[0];
					var itemids=new2str[k].split("^")[1];
					var itemcss='txt';
					
					var iiarr=itemvalue.split('|')
					var vv='';
					if (iiarr[1]!=''){
						vv="<img src='" + iiarr[1] +"' width='25' height='25' title='" + iiarr[0] +"'/>";
					}else{
					    vv=iiarr[0];
					}
					if (num==2){
					strs+='<span id="att'+num+k+'" class="'+itemcss+'" onclick="shop.getAttr(this,2,'+a+','+vlan+',3,\''+itemids+'\')">'+vv+'<i></i></span> ';
					}else{
					strs+='<span id="att'+num+k+'" class="'+itemcss+'" onclick="shop.getAttr(this,3,'+a+','+vlan+',0,'+itemids+')">'+vv+'<i></i></span> ';
					}
				}
				//alert(strs);
			    jQuery("#showattr"+num).html(strs);
               // alert(str22);
			  }
			  var attr='';
			  var attrs='';
			  for(var k=0;k<=a;k++)
			  {
			   if ($("#attr"+k).val()!='' && $("#attr"+k).val()!=undefined){
			    if (attr==''){
			      attr='“'+$("#attr"+k).val()+'”';
				  attrs=$("#attrname"+k).html()+$("#attr"+k).val();
			     }else{
			      attr+=',“'+$("#attr"+k).val()+'”';
				  attrs+='  '+$("#attrname"+k).html()+$("#attr"+k).val();
				 }
			   }
			  }
			  $("#AttributeCart").val(attrs);
			  $("#buyselect").html("<b>已选择：<font color=brown>"+attr+"</font></b>");
			 },
			 addCart:function(ev,id,a){
			  if (shop.check(a)){
			    str="loading...";
				var kbid='';
				$("input[name=Bundid]").each(function(){
					if ($(this)[0].checked){
						kbid+=','+$(this)[0].value;
					}
				 });
				
				var top=$('#carbtn').offset().top - $(document).scrollTop()+28;
				if (top<0) top=28;
				var left=parseInt($('#carbtn').offset().left);
				if (left+400-parseInt($(document.body).width())>0) left=left-300;
				box=$.dialog({id:'mycart',title:'购物车内商品',max:false,min:false,content:str,width:400,height:150,left: left,top: top,init: function(){
				jQuery.getScript(dir+'plus/ajaxs.asp?kbid='+kbid+'&id='+id+'&attrid='+$("#attrid").val()+'&action=addCart&num='+$("#num").val()+'&AttributeCart='+escape($("#AttributeCart").val())+'&istype='+$("#istype").val(),function(){
				  if (data.flag=='error'){
						    alert('对不起，您没有登录!');
					}else if (data.flag=='error1'){
						    alert('对不起，您所在的用户级别不能购买本商品!');
					   }else{
				            box.content("<img src='"+dir+"shop/images/suc.gif' align='absmiddle'><span style='font-size:14px;color:#000;'>已成功添加到购物车！</span><br/><form name='paymentform' action='"+dir+"shop/payment.asp' method='post'><div id='cartShow' style='height:140px'>"+unescape(data.str)+"</div><div class='jrgwc'><input type='image' src=\'"+dir+"shop/images/hesuan.gif\'> <a href=\'"+dir+"shop/shoppingcart.asp\'><img src=\'"+dir+"shop/images/chakangouwuche.gif\'></a></form></div>");
					   }																														                 
				 });
						
						}
					});
				
				
			   }
			 }
 }
 //-->

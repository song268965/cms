myFocus.pattern.extend({
	'mF_classicHB':function(settings,$){//*********************缁忓吀镐€镞х郴鍒椾簩--娴锋姤椋庢牸******************
		var $focus=$(settings);
		var $picList=$focus.find('.pic li');
		var $txtList=$focus.addListTxt().find('li');
		var $numList=$focus.addListNum().find('li');
		//CSS
		var w=settings.width,h=settings.height,txtH=settings.txtHeight;//璁剧疆榛樿镄勬枃瀛楅佩搴?
		for(var i=0,n=$picList.length;i<n;i++){
			$picList.eq(i).css({display:'none',top:-0.1*h,left:-0.1*w,width:1.2*w,height:1.2*h});
			$txtList.eq(i).css({top:-txtH});
		}
		//PLAY
		$focus.play(function(i){
			$picList.eq(i).stop().css({display:'none',top:-0.1*h,left:-0.1*w,width:1.2*w,height:1.2*h});
			$txtList.eq(i).stop().css({top:-txtH});
			$numList.eq(i).slide({width:16},200)[0].className='';
		},function(i){
			$picList.eq(i).fadeIn(300).slide({width:w,height:h,top:0,left:0},300);
			$txtList.eq(i).slide({top:0},300);
			$numList.eq(i).slide({width:26},200)[0].className='current';
		});
		//Control
		$focus.bindControl($numList);
		if (settings.wrap=='false') jQuery('#'+settings.id).parent().removeClass(settings.pattern+'_wrap');
	}
});
myFocus.config.extend({'mF_classicHB':{txtHeight:20}});//榛樿鏂囧瓧灞傞佩搴
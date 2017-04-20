myFocus.pattern.extend({//*********************鐧惧彾绐?*****************
	'mF_shutters':function(settings,$){
		var c=Math.floor(settings.width/50);
		var $focus=$(settings);
		var $txtList=$focus.addListTxt().find('li');
		var $picUls=$focus.find('.pic ul').repeat(c);
		var $prevBtn=$focus.addHtml('<div class="prev"><a href="javascript:;">PREV</a></div>');
		var $nextBtn=$focus.addHtml('<div class="next"><a href="javascript:;">NEXT</a></div>');
		var $picListArr=[];
		//CSS
		var w=settings.width/c;
		$picUls.each(function(i){
			$(this).css({width:w*(i+1),zIndex:c-i});
			$picListArr.push($(this).find('li'));
		});
		//PLAY
		var running=false,done=0;//璁板綍杩愯钟舵€?
		$focus.play(function(i){
			running=true;
			$txtList[i].className='';
		},function(i){
			$txtList[i].className='current';
			for(var j=0;j<c;j++) timeoutFx($picListArr[j],i,(j+1)*100);//姣忓垏鐗囧欢镞?00姣
		});
		function timeoutFx($picList,i,t){
			setTimeout(function(){
				$picList.eq(i).css({zIndex:1}).fadeIn(function(){
					$picList.each(function(){this.style.display='none'});
					this.style.cssText='z-index:"";display:block';
					done+=1;
					if(done===c) running=false,done=0;
				});
			},t);
		}
		//Control
		$prevBtn.bind('click',function(){if(!running) $focus.run('-=1')});
		$nextBtn.bind('click',function(){if(!running) $focus.run('+=1')});
		$focus.bind('mouseover',function(){$prevBtn.addClass('hover'),$nextBtn.addClass('hover')});
		$focus.bind('mouseout',function(){$prevBtn.removeClass('hover'),$nextBtn.removeClass('hover')});
	}
});
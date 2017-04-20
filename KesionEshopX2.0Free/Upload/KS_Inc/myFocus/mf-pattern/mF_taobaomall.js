myFocus.pattern.extend({//*********************娣桦疂鍟嗗煄椋庢牸******************
	'mF_taobaomall':function(settings,$){
		var $focus=$(settings);
		var $picBox=$focus.find('.pic');
		var $picUl=$picBox.find('ul');
		var $txtList=$focus.addListTxt().find('li');
		$picUl[0].innerHTML+=$picUl[0].innerHTML;//镞犵绅澶嶅埗
		//CSS
		var n=$txtList.length,dir=settings.direction,prop=dir==='left'?'width':'height',dis=settings[prop];
		$picUl.addClass(dir)[0].style[prop]=dis*n*2+'px';
		$picUl.find('li').each(function(){//娑堥櫎涓娄笅li闂寸殑澶氢綑闂撮殭
			$(this).css({width:settings.width,height:settings.height});
		});
		var txtH=settings.txtHeight;
		$focus.css({height:settings.height+txtH+1});
		$picBox.css({width:settings.width,height:settings.height});
		$txtList.each(function(){this.style.width=(settings.width-n-1)/n+1+'px'});
		$txtList[n-1].style.border=0;
		//PLAY
		var param={};
		$focus.play(function(i){
			$txtList[i>=n?(i-n):i].className = '';
		},function(i){
			param[dir]=-dis*i;
			$picUl.slide(param,settings.duration,settings.easing);
			$txtList[i>=n?(i-n):i].className = 'current';
		},settings.seamless);
		//Control
		$focus.bindControl($txtList);
	}
});
myFocus.config.extend({
	'mF_taobaomall':{//鍙€変釜镐у弬鏁?
		txtHeight:28,//榛樿镙囬鎸夐挳楂桦害
		seamless:true,//鏄惁镞犵绅锛歵rue(鏄?| false(鍚?
		duration:600,//杩囨浮镞堕棿(姣)锛屾椂闂磋秺澶ч€熷害瓒婂皬
		direction:'top',//杩愬姩鏂瑰悜锛屽彲阃夛细'top'(鍚戜笂) | 'bottom'(鍚戜笅) | 'left'(鍚戝乏) | 'right'(鍚戝彸)
		easing:'easeOut'//杩愬姩褰㈠纺锛屽彲阃夛细'easeOut'(蹇嚭鎱㈠叆) | 'easeIn'(鎱㈠嚭蹇叆) | 'easeInOut'(鎱㈠嚭鎱㈠叆) | 'swing'(鎽囨憜杩愬姩) | 'linear'(鍖€阃熻繍锷?
	}
});
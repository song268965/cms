myFocus.pattern.extend({//*********************娣桦疂2010涓婚〉椋庢牸******************
	'mF_taobao2010':function(settings,$){
		var $focus=$(settings);
		var $picUl=$focus.find('.pic ul');
		var $txtList=$focus.addListTxt().find('li');
		var $numList=$focus.addListNum().find('li');
		$picUl[0].innerHTML+=$picUl[0].innerHTML;//镞犵绅澶嶅埗
		//CSS
		var n=$txtList.length,dir=settings.direction,prop=dir==='left'?'width':'height',dis=settings[prop];
		$picUl.addClass(dir)[0].style[prop]=dis*n*2+'px';
		$picUl.find('li').each(function(){//娑堥櫎涓娄笅li闂寸殑澶氢綑闂撮殭
			$(this).css({width:settings.width,height:settings.height});
		});
		//PLAY
		var param={};
		$focus.play(function(i){
			var index=i>=n?(i-n):i;
			$numList[index].className = '';
			$txtList[index].style.display = 'none';
		},function(i){
			var index=i>=n?(i-n):i;
			param[dir]=-dis*i;
			$picUl.slide(param,settings.duration,settings.easing);
			$numList[index].className = 'current';
			$txtList[index].style.display = 'block';
		},settings.seamless);
		//Control
		$focus.bindControl($numList);
	}
});
myFocus.config.extend({
	'mF_taobao2010':{//鍙€変釜镐у弬鏁?
		seamless:true,//鏄惁镞犵绅锛屽彲阃夛细true(鏄?/false(鍚?
		duration:600,//杩囨浮镞堕棿(姣)锛屾椂闂磋秺澶ч€熷害瓒婂皬
		direction:'left',//杩愬姩鏂瑰悜锛屽彲阃夛细'top'(鍚戜笂) | 'left'(鍚戝乏)
		easing:'easeOut'//杩愬姩褰㈠纺锛屽彲阃夛细'easeOut'(蹇嚭鎱㈠叆) | 'easeIn'(鎱㈠嚭蹇叆) | 'easeInOut'(鎱㈠嚭鎱㈠叆) | 'swing'(鎽囨憜杩愬姩) | 'linear'(鍖€阃熻繍锷?
	}
});
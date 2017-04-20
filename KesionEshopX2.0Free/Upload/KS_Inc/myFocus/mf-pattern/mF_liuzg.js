/*
* 缁氢附鍒囩墖椋庢牸 v2.0
* Date 2012.5.25
* Author koen_lee
*/
myFocus.pattern.extend({
	'mF_liuzg':function(settings,$){
		var $focus=$(settings);
		var $picBox=$focus.find('.pic');
		var $picList=$picBox.find('li');
		var $txtList=$focus.addListTxt().find('li');
		var $numList=$focus.addListNum().find('li');
		//HTML++
		var c=Math.floor(settings.height/settings.chipHeight),n=$txtList.length,html=['<ul>'];
		for(var i=0;i<c;i++){
			html.push('<li><div>');
			for(var j=0;j<n;j++) html.push($picList[j].innerHTML);
			html.push('</div></li>');
		}
		html.push('</ul>');
		$picBox[0].innerHTML=html.join('');
		//CSS
		var w=settings.width,h=settings.height,cH=Math.round(h/c);
		var $picList=$picBox.find('li'),$picDivList=$picBox.find('div');
		$picList.each(function(i){
			$picList.eq(i).css({width:w,height:cH});
			$picDivList.eq(i).css({height:h*n,top:-i*h});
		});
		$picBox.find('a').each(function(){this.style.height=h+'px'});
		//PLAY
		$focus.play(function(i){
			$txtList[i].style.display='none';
			$numList[i].className='';
		},function(i){
			var tt=settings.type||Math.round(1+Math.random()*2);//鏁堟灉阃夋嫨
			var dur=tt===1?1200:300;
			for(var j=0;j<c;j++){
				$picDivList.eq(j).slide({top:-i*h-j*cH},tt===3?Math.round(300+(Math.random()*(1200-300))):dur);
				dur=tt===1?dur-150:dur+150;
			}
			$txtList[i].style.display='block';
			$numList[i].className = 'current';
		});
		//Control
		$focus.bindControl($numList);
	}
});
myFocus.config.extend({
	'mF_liuzg':{//鍙€変釜镐у弬鏁?
		chipHeight:36,//锲剧墖鍒囩墖楂桦害(镀忕礌)锛岃秺澶у垏鐗囧瘑搴﹁秺灏?
		type:0////鍒囩墖鏁堟灉锛屽彲阃夛细1(鐢╁ご) | 2(鐢╁熬) | 3(鍑屼贡) | 0(闅忔満)
	}
});
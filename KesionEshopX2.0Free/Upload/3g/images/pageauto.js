// JavaScript Document
	$(function(){
		auto();
	});
	$(window).resize(function(){
		auto();

	});
	function auto(){
		var meta = $("body").width()
		var a = 375/20;
		var d = meta/a;
		if(meta>375){
			$("html").css({"font-size":""+d+"px"});
		}
		else{
			$("html").css({"font-size":"20px"});
		};
	};


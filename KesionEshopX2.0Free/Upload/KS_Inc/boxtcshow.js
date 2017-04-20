var pgg_s = function(settings) {
	var defaults = {
	gco_html:"",	
	ggtype:"top",
	pic_s:"0px",
	picgl:"300px",
	s_opacity: "0",//透明度
	e_opacity: "10",//透明度
	timing: "500"//动画时间
	};
	var settings = $.extend(defaults, settings);
				picgg_box=$.dialog({
				id: 'Tips',
				title: "订单处理",
				left: '100%',
       			top: '100%',
				height:"60px",
				content: settings.gco_html,
				fixed: true,
				min:false,
				max:false,
				padding: '0px',
				resize: false,
				init: function(){
						var duration = 300, /*动画时长*/
						api = this,
						opt = api.config,
						wrap = api.DOM.wrap;
						wrap.css(settings.ggtype, settings.pic_s);
						wrap.css('opacity', settings.s_opacity);
						switch (settings.ggtype)
						{
						case "top":
						  wrap.animate({top:settings.picgl, opacity:settings.e_opacity}, settings.timing, function(){});
						  break;
						case "left":
						  wrap.animate({left:settings.picgl, opacity:settings.e_opacity}, settings.timing, function(){});
						  break;
						case "right":
						  wrap.animate({right:settings.picgl, opacity:settings.e_opacity}, settings.timing, function(){});
						  break;
						case "bottom":
						  wrap.animate({bottom:settings.picgl, opacity:settings.e_opacity}, settings.timing, function(){});
						  break;
						default:
						  wrap.animate({top:settings.picgl, opacity:settings.e_opacity}, settings.timing, function(){});
						}
						
				},
				close:function(){
						var duration = 300, /*动画时长*/
							api = this,
							opt = api.config,
							wrap = api.DOM.wrap;
						wrap.animate({top:'-'+settings.picgl,opacity:0}, settings.timing, function(){
							opt.close = function(){};
							api.close();
						});
						return false;
					}
				});
	
	
}

function pic_move(obj,fx){
				var box_d=$(obj).parent().parent().parent();
				var box_fx='';
				if (fx==1)
				{box_fx=box_d.prev()}
				else if(fx==2)
				{box_fx=box_d.next()}
				else
				{box_fx=''}
				if ( box_fx.length > 0 ) {
					if(fx==1) box_fx.before(box_d.clone())
					if(fx==2) box_fx.after(box_d.clone())
					if(fx==1 || fx==2)box_d.remove()
				} 
}


$.dialog.notice = function( options )
{
    var opts = options || {},
        api, aConfig, hide, wrap, top,
        duration = opts.duration || 800;
        
    var config = {
        id: 'Notice',
        left: '100%',
        top: '100%',
        fixed: true,
        drag: false,
        resize: false,
		min:false,
		max:false,
        init: function(here){
            api = this;
            aConfig = api.config;
            wrap = api.DOM.wrap;
            top = parseInt(wrap[0].style.top);
            hide = top + wrap[0].offsetHeight;
                        
            wrap.css('top', hide + 'px')
            .animate({top: top + 'px'}, duration, function(){
                opts.init && opts.init.call(api, here);
            });
        },
        close: function(here){
            wrap.animate({top: hide + 'px'}, duration, function(){
                opts.close && opts.close.call(this, here);
                aConfig.close = $.noop;
                api.close();
            });
                        
            return false;
        }
    };
        
    for(var i in opts)
    {
        if( config[i] === undefined ) config[i] = opts[i];
    }
        
    return $.dialog( config );
};



function stopInterval(){
    clearInterval(checkInterval);
}



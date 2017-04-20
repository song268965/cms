document.writeln("<div id=\"float\" style=\"Z-INDEX: 99; FILTER: ALPHA(opacity=80); LEFT: 159px; WIDTH: 142px; POSITION: absolute; TOP: 400px\" align=\"center\"><input type=\"image\" onClick=\"comchk();\" name=\"imgbtnFloat\" id=\"imgbtnFloat\" src=\"/shop/images/compare_float.gif\" alt=\"\" border=\"0\" /><br>");
document.writeln("<DIV id=com_item align=\"center\"></DIV>");
document.writeln("</DIV>");

var sel = new Array();
			var sel_num = 0;
			
			function getCookieVal (offset) 
			{
				var endstr = document.cookie.indexOf (";", offset);
  				if (endstr == -1)
    				endstr = document.cookie.length;
  				return unescape(document.cookie.substring(offset, endstr));

			}
			function GetCookie (name) 
			{
  				var arg = name + "=";
  				var alen = arg.length;
  				var clen = document.cookie.length;
  				//alert(document.cookie.length);
  				var i = 0;
  				while (i < clen) 
  				{
    					var j = i + alen;
    					if (document.cookie.substring(i, j) == arg)
      						return getCookieVal (j);
    					i = document.cookie.indexOf(" ", i) + 1;
    					if (i == 0) break; 
  				}
  				return null;
			}
			function SetCookie (name,value,expires,path,domain,secure) 
			{
  				document.cookie = name + "=" + escape (value) +
    				((expires) ? "; expires=" + expires.toGMTString() : "") +
    				((path) ? "; path=" + path : "") +
    				((domain) ? "; domain=" + domain : "") +
    				((secure) ? "; secure" : "");
					return value;
			}
			function DeleteCookie (name)
			{
				if(GetCookie(name) != null)
				{
					SetCookie(name,"");
				}
			}
			//Cookie

			function cookie_content()
			{
				i = 0;
				var content = "";
				for(key in sel)
				{
					if(i == 0)
					{
						content += key + "[" + sel[key] + "]";
					}
					else
					{
						content += "," + key + "[" + sel[key] + "]";
					}
					i++;
				}
				//alert(content);
				return content;
			}
			function add(id,nm) {
				if(!sel[id])
				{
					if(sel_num >= 5) 
					{
					alert('最多只能选择5个商品进行比较!');
					}
					else 
					{
						sel_num++;
						sel[id] = nm;
					}
				}
				else 
				{
					sel2 = new Array();
					for(key in sel) 
					{
						if(id!=key) 
						{
							sel2[key] = sel[key];
						}
					}
					sel = sel2;
					sel_num--;
				}
				SetCookie("PRODUCT_COMPARE_COOKIE",cookie_content(),null,"/",null);
				//alert(GetCookie("PRODUCT_COMPARE_COOKIE"));
				draw();
			}
			
			function del(id) {
			    if (!confirm('确定删除吗？')) return;
				sel2 = new Array();
				for(key in sel) 
				{
					if(id!=key) 
					{
						sel2[key] = sel[key];
					}
				}
				sel = sel2;
				sel_num--;
				SetCookie("PRODUCT_COMPARE_COOKIE",cookie_content(),null,"/",null);
				draw();
			}
			function draw() {
				out = '';
				for(key in sel) {
					out += "<font color=#ff7312 size=1><b>|</b></font><br><input type=button onclick=\"del('"+key+"')\" value='"+sel[key]+"'  style='border:1px solid #ff7312 ;background-color:#ffffff;height:24;width:120px;cursor:pointer;color:'black';'><br>";
				}
				out+="<br>"
				com_item.innerHTML = out;
			}
			function comchk() {
				if(sel_num < 2) {
					alert('请至少选择两个商品!')
				}else {
					out = '';
					i=0;
					for(key in sel){
						++i;
						out += "ids=" ;
						out +=key+'&';
						
					}
					var str = "/shop/compare.asp?"+out;
					window.open(str);
				}
			}
			function inni_data()
			{
				//alert(GetCookie("PRODUCT_COMPARE_COOKIE"));
				var cookie_sel = new Array();
				cookie_str = GetCookie("PRODUCT_COMPARE_COOKIE");
				if(cookie_str != "" && cookie_str != null)
				{
					cookie_sel = cookie_str.split(",");
					for( key in cookie_sel)
					{
						//alert(cookie_sel[key]);
						i = cookie_sel[key].indexOf("[");
						j = cookie_sel[key].indexOf("]");
						//alert(cookie_sel[key].substring(0, i));
						//alert(cookie_sel[key].substring(i+1, j));
						sel[cookie_sel[key].substring(0, i)] = cookie_sel[key].substring(i+1, j);
						sel_num++;
					}
					draw();
				}
			}

inni_data();

lastScrollY = 0;

function heartBeat() {
	document.getElementById("float").style.left=30;
	diffY = document.body.scrollTop;
	 if (diffY==0) diffY=document.documentElement.scrollTop;

	
	percent =.2*(diffY-lastScrollY);

	if(percent>0){
		percent = Math.ceil(percent);
	}else{
		percent = Math.floor(percent);
	}
	
	document.getElementById("float").style.pixelTop+= percent;
	lastScrollY = lastScrollY+percent;
}
window.setInterval("heartBeat()",1);
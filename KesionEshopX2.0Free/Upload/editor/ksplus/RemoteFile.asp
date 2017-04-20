<!DOCTYPE html>
<html>
<head >
    <script src="../../KS_Inc/Jquery.js" type="text/javascript"></script>
    <script src="../../KS_Inc/Common.js" type="text/javascript"></script>
    <style>
	 BODY   {border: 0; margin: 0; cursor: default;font-family:宋体; font-size:9pt;}
	.textbox{ padding:3px; border:1px solid; border-color:#666 #ccc #ccc #666; background:#F9F9F9; color:#333; resize: none;}
	.textbox:hover, .textbox:focus, textarea:hover, textarea:focus{ border-color:#09C; background:#F5F9FD; }
	</style>
</head>
<body scroll="auto">

 <script type="text/javascript">
     function ok() {
         if (document.getElementById('collecthttp').value == '') {
             KesionJS.Alert("请输入远程图片地址,一行一张地址！", "document.getElementById('collecthttp').focus();");
             return false;
         } else {
		    if (top.frames['MainFrame']==undefined){
			 top.ProcessCollect(document.getElementById('collecthttp').value);
			}else{
             top.frames['MainFrame'].ProcessCollect(document.getElementById('collecthttp').value);
			}
			
         }
     };
	</script>
<div style='padding:3px'>带http://开头的远程图片地址,每行一张图片地址:<br/>
<textarea id='collecthttp' style='width:95%;height:150px' class='textbox'></textarea>
<div style="text-align:center">
<input type="button" value=" 确 定 " class="button" onclick="ok()"/>
</div>
</div>
</body>
</html>
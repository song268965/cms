 /*
 作用:使MainFrame始终获得焦点
 */
 document.onkeydown=keyDown;
 function keyDown()
 { // alert(typeof(parent.frames['MainFrame'])); 
  if ((event.ctrlKey)||(event.altKey)||(event.keyCode==46))   
   {
     if (typeof(parent.frames['MainFrame'])=='object')
     { 
      parent.frames['MainFrame'].focus();
	  }
    else
     top.frames['MainFrame'].focus();	
  }
 }
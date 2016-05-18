function killErrors(){return true;}
window.onerror = killErrors;
var vcome = escape(document.referrer);
var vpage = escape(window.location.href);
var screeninfo1 = escape(window.screen.width);
var screeninfo2 = escape(window.screen.height);
var getpage = '../Tjserver/Tj.asp?vcome='+vcome+'&vpage='+vpage+'&screeninfo1='+screeninfo1+'&screeninfo2='+screeninfo2;
document.write('<script src='+getpage+'></script>');
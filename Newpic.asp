<!--#include file="Inc/SysProduct.asp" -->
<% 
set rs=server.createobject("adodb.recordset")
sql="select * from newpic"
rs.open sql,conn,1,3
pic1=rs("pic1")
pic2=rs("pic2")
pic3=rs("pic3")
pic4=rs("pic4")
pic5=rs("pic5")
pic6=rs("pic6")
movie1=rs("movie1")
movie2=rs("movie2")
movie3=rs("movie3")
movie4=rs("movie4")
movie5=rs("movie5")
movie6=rs("movie6")
rs.close
set rs=nothing
%>

							<SCRIPT type=text/javascript>
var img=new Array();
var txt=new Array();
var lnk=new Array();

img[0]='<%=pic1%>'
lnk[0]=escape('');
txt[0]='';

img[1]='<%=pic2%>'
lnk[1]=escape('');
txt[1]='';

img[2]='<%=pic3%>'
lnk[2]=escape('');
txt[2]='';

img[3]='<%=pic4%>'
lnk[3]=escape('');
txt[3]='';

img[4]='<%=pic5%>'
lnk[4]=escape('');
txt[4]='';


 var focus_width=261;
 var focus_height=193;
 var text_height=0;



var swf_height = focus_height+text_height;

var pics=img.join("|");
var links=lnk.join("|");
var texts=txt.join("|");
 
 document.write('<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="[url]http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6[/url],0,0,0" width=259 height=191>');
document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="flash/pixviewer.swf"><param name="quality" value="high"><param name="bgcolor" value="#ffffff">');
document.write('<param name="menu" value="false"/><param name="wmode" value="transparent"/>');
document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
document.write('<embed src="flash/pixviewer.swf" wmode="transparent" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#ffffff" quality="high" width="'+ focus_width +'" height="'+ focus_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash"  pluginspage="[url]http://www.macromedia.com/go/getflashplayer[/url]" />');
document.write('</object>');    </SCRIPT>
<table width="992" border="0" cellpadding="0" cellspacing="0" align="center">
  <!--DWLayoutTable-->
  <tr>
    <td height="77" colspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="Images/top_back.jpg">
      <!--DWLayoutTable-->
      <tr>
        <td width="259" rowspan="2" valign="middle"><div align="center"><img src="<%=LogoUrl%>"></div></td>
        <td height="19" colspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <!--DWLayoutTable-->
          <tr>
            <td width="737" height="19" valign="top" background="Images/top_back1.gif" align="right"><a href="#"><span class="STYLE1">公司邮局</span></a> <a href="Aboutus.asp?Title=联系我们"><span class="STYLE1">联系我们</span></a> <a href="#" onClick="window.external.addFavorite('<%=SiteUrl%>','<%=SiteName%>')"><span class="STYLE1">收藏本站</span></a> <a id="based" style="color:red"><span class="STYLE1">繁体中文</span></a><Script Language=Javascript Src="Inc/southidcj2f.Js"></Script> <a href="enindex.asp"><span class="STYLE1">English&nbsp;&nbsp; </span></a></td>
          </tr>
        </table>        </td>
      </tr>
      <tr>
        <td width="426" height="58" valign="top"><div align="center"><% If IsFlash="Yes" Then %>
	<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="426" height="58">
        <param name="movie" value="<%=BannerUrl%>">
        <param name="quality" value="high">
        <param name="wmode" value="transparent">
        <embed src="<%=BannerUrl%>" width="426" height="58" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" wmode="transparent"></embed></object>
		<% Else %>
		  <img src="<%=BannerUrl%>" width="426" height="58"> 
        <% end If%></div></td>
        <td width="315" valign="bottom" align="right"><table cellSpacing=0 cellPadding=0 width="100%" border=0>
        <form method="Get" name="myform" action="search.asp" target="_blank">
          <tr>
            
            <td align="right" width="230"><input type="text" name="keyword"  size=12 value="关键字" maxlength="50" onFocus="this.select();"> 
<select name="Field" size="1">
          <option value="Title" selected>产品名称</option>
          <option value="Content">产品说明</option>
        </select> 
</td>
            <td width="65" align="right"><input name=Submit type=image id=Submit src='images/search.gif' style="border: 0px;width:70px;height:24px;">
              <input id=Field type=hidden value=Title name=Field></td>
          </tr>
        </FORM>
    </table> </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td width="194" rowspan="2" valign="bottom" background="Images/top_back2.gif"> 
                    <script language="JavaScript">
today=new Date();
function initArray(){
this.length=initArray.arguments.length
for(var i=0;i<this.length;i++)
this[i+1]=initArray.arguments[i]  }
var d=new initArray(
" 星期日",
" 星期一",
" 星期二",
" 星期三",
" 星期四",
" 星期五",
" 星期六");
document.write(
"<font color=#FFFFFF style='font-size:11px;font-family: verdana'> ",
today.getYear(),"年",
today.getMonth()+1,"月",
today.getDate(),"日",
d[today.getDay()+1],
"</font>" ); 
					</script></td>
    <td width="806" height="37" background="Skin/11/m-bg1-2.gif" valign="bottom">
              <% if ChannelID<>0 then
	  call ShowRootClass_menu(1)
	  end if %>
    </td>
  </tr>
  <tr>
    <td height="1"></td>
  </tr>
  <tr>
    <td height="2"></td>
    <td></td>
  </tr>
</table>
<table width="900" height="69" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><%=Copyright%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>

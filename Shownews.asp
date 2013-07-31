<!--#include file="Inc/SysProduct.asp" -->
<!-- #include file="Head.asp" -->
<%owen=request("id")%>
<script language="JavaScript" type="text/JavaScript">
function fontZoom(size){
 document.getElementById('fontZoom').style.fontSize=size+'px'
}
</SCRIPT>
<SCRIPT language=JavaScript>
var currentpos,timer;

function initialize()
{
timer=setInterval("scrollwindow()",50);
}
function sc(){
clearInterval(timer);
}
function scrollwindow()
{
currentpos=document.body.scrollTop;
window.scroll(0,++currentpos);
if (currentpos != document.body.scrollTop)
sc();
}
document.onmousedown=sc
document.ondblclick=initialize
</SCRIPT>
<% 
id=cstr(request("id"))
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
sql="update news set hits=hits+1 where id="&id
conn.execute sql
sql="select * from news where id="&owen
rsnews.Open sql,conn,1,1
title=rsnews("title")
if rsnews.eof and rsnews.bof then
response.Write("数据库出错")
else
%>
<body topmargin="0">

<DIV id=index_container>
  <div id="index_left">
    <div id="left">
      <div id="left_title">新闻中心</div>
      <div id="left_content">
        <DIV id=elem-FrontProductCategory_slideTree-001 name="商品分类"><LINK 
href="imgbyy/css2.css" type=text/css 
rel=stylesheet>

<DIV class=FrontProductCategory_slideTree-default 
id=FrontProductCategory_slideTree-001>
<DIV class=left_bottom>
<DIV class=right_bottom>
<DIV class=left_top>
<DIV class=right_top>


<UL id=treemenu1>

<%
Set rslist = Server.CreateObject("ADODB.Recordset")
sql="select BigClassName from BigClass_New order by BigClassID"
rslist.open sql,conn,1,3
do while not rslist.eof
%>

  <LI>
<a href="NewsClass.asp?BigClass=<%=rslist("BigClassName")%>"><%=rslist("BigClassName")%></a>  
  </li>
  
  
  
<%rslist.movenext 
loop
rslist.close
set rslist=nothing
%>


  </UL>
 <table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10"></td>
  </tr>
</table>
 
  
  
</DIV>


</DIV></DIV></DIV></DIV>
</DIV>
      </div>
      <div id="left_btm"></div>
    </div>
	<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="6"></td>
  </tr>
</table>

    <div id="index_search">
      <div id="index_search1"></div>
      <div id="index_search2">
        <div class="query">
          <table width="95%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="5"></td>
            </tr>
          </table>
          <% call ShowSearch(1) %>
          <table width="95%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="5"></td>
            </tr>
          </table>
          <table width="95%" height="65" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="30" align="center"><a onFocus="this.blur()" 
href="msnim:chat?contact=59684111@qq.com"><img 
src="imgbyy/msn.png" alt="MSN" width="158" height="50" border="0" /></a></td>
            </tr>
            <tr>
              <td height="5"></td>
            </tr>
            <tr>
              <td height="30"></td>
            </tr>
          </table>
        </div>
      </div>
      <div id="index_search3"></div>
    </div>
    <div id="index_contact">
      <div id="index_contact1"></div>
      <div id="index_contact2">
        <div id="elem-FrontSpecify_show-001" name="联系我们说明页">
          <link 
href="about_files/css3.css" type="text/css" rel="stylesheet" />
          <div class="FrontSpecify_show-default" id="FrontSpecify_show-001">
            <div class="content">
              <div class="style1"><font face="Arial">地址：<font face="Arial"><%=Indexadd%><font 
face="Arial"></font><br />
                </font></font>电话：<%=Indextel%> <br />
                传真：<%=Indexfax%><font face="Arial"></font> </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
  <DIV id=index_right><!--<div id="right_title">公司简介</div> -->
<DIV id=right_content>
<DIV ><LINK 
href="about_files/FrontCommonContent_showDetailByList.css" type=text/css 
rel=stylesheet>
<DIV class=FrontCommonContent_showDetailByList-default 
id=>
<DIV class=content>
<DIV id=right_title>信息详情</DIV>
<DIV class=text>
<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10"></td>
  </tr>
</table>

  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!--DWLayoutTable-->
    <tr>
      <td width="804" valign="top"><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="50" align="center" class="tit"><%= rsnews("title") %></td>
          </tr>
          <tr align="center">
            <td width="80%" height="30"  style="BORDER-RIGHT: #e9e9e9 1px solid; BORDER-TOP: #e9e9e9 1px solid; BORDER-LEFT: #e9e9e9 1px solid; BORDER-BOTTOM: #e9e9e9 1px solid" >发布者：<%= rsnews("user") %> 发布时间：<%= rsnews("AddDate") %> 阅读：<font color="#FF0000"><%= rsnews("hits") %></font>次 【字体：<a class="black" href="javascript:fontZoom(16)">大</a> <a class="black" href="javascript:fontZoom(14)">中</a> <a class="black" href="javascript:fontZoom(12)">小</a>】</td>
          </tr>
          <tr>
            <td class="black" id="fontzoom"><br />
                <%=rsnews("content") %></td>
          </tr>
          <tr align="right">
            <td>&nbsp;</td>
          </tr>
          <tr align="right">
            <td>&nbsp;</td>
          </tr>
          <% 
		end if
		rsnews.close
		set rsnews=nothing
		%>
        </table>
          <% if NewsComment="Yes" then  %>
          <table bordercolor="#d8d8f0" cellspacing="0" cellpadding="0" width="95%" align="center" >
            <tr>
              <td width="800" height="22" align="middle" class="tr"><p align="left"> <b>　</b><b>相关评论</b></p></td>
            </tr>
            <tr>
              <td align="middle" bgcolor="#ffffff" height="248" style="padding-top: 5px; padding-bottom: 5px"><%set rs3 = server.CreateObject("adodb.recordset")
	  sql="select * from comment where com_typeid="&id&" order by com_id desc"
	  rs3.open sql,conn,1,1
	   if rs3.bof and rs3.eof then
         %>
                  <table width="98%" border="0" cellpadding="0" cellspacing="0" bordercolor="#d8d8f0" >
                    <tr>
                      <td height="22">暂无评论</td>
                    </tr>
                  </table>
                <br />
                  <%else
         do while not rs3.eof         
%>
                  <table width="98%" height="46" border="0" cellpadding="0" cellspacing="0" bordercolor="#d8d8f0" id="table13">
                    <tr>
                      <td width="100%" height="22" class="tr"><span lang="en" xml:lang="en">&nbsp;&nbsp; </span>评论人：<%=Replace(Replace(rs3("com_name"),"<","&lt;"),">","&gt;")%>&nbsp; 评论时间：<%=rs3("com_date")%></td>
                    </tr>
                    <tr>
                      <td height="22" style="word-break:break-all"><%response.write "&nbsp;&nbsp;&nbsp;&nbsp;" & rs3("com_content") & "<br>"%>
                      </td>
                    </tr>
                  </table>
                <br />
                  <%rs3.movenext
		loop
		end if
		rs3.close%>
                  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#e9e9e9" id="table12">
                    <form action="Comment.asp?id=<%=id%>" method="post" name="svcomment" id="svcomment">
                      <tr bgcolor="#FFFFFF">
                        <td height="21" colspan="3">&nbsp;<a name="comment" id="comment">发表评论</a>：</td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td width="15%" height="20">&nbsp;呢称：</td>
                        <th width="85%" height="20"> <div align="left">
                            <input type="text" name="com_name" size="30" />
                        </div></th>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td height="20" >&nbsp;评论内容：</td>
                        <th height="20"><div align="left">
                            <textarea rows="7" name="com_content" cols="50"></textarea>
                        </div></th>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <th height="33" colspan="3"> <div align="center">
                            <input type="submit" value="提交评论" name="B1" />
                          
                          <input type="reset" value="重置" name="B2" />
                        </div></th>
                      </tr>
                    </form>
                  </table></td>
            </tr>
          </table>
        <% end if %>
          <br />
          <table width="95%" height="30" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td align="right"><a href="javascript:window.print()">打印本页</a> || <a href="javascript:window.close()">关闭窗口</a></td>
            </tr>
        </table></td>
    </tr>
  </table>
</DIV>
</DIV></DIV></DIV></DIV>
<DIV class=clear></DIV></DIV>
<DIV class=clear></DIV>
</DIV></DIV>
<DIV class=clear></DIV>

<DIV id=footer>
<DIV id=elem-FrontSpecify_show-020 ><LINK 
href="imgbyy/css3.css" type=text/css rel=stylesheet>
<DIV class=FrontSpecify_show-default id=FrontSpecify_show-020>
<DIV class=content>
<!--#include file="Inc/Foot.asp" -->


</DIV></DIV></DIV></DIV>


</BODY></HTML>

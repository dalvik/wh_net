<!--#include file="Inc/SysProduct.asp" -->

<!-- #include file="Head.asp" -->
<% 
BigClass=request("BigClass")
SmallClass=request("SmallClass")
%><body topmargin="0">

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
<DIV id=right_title><%=BigClass%></DIV>
<DIV class=text>
<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10"></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <% 
page=clng(request("page"))		 
Set rs=Server.CreateObject("ADODB.RecordSet") 
if BigClass<>"" and SmallClass <>"" then
sql="select * from news where BigClassName='"&BigClass&"' and SmallClassName='"&SmallClass&"' order by AddDate desc"
rs.Open sql,conn,1,1
elseif BigClass<>"" then
sql="select * from news where BigClassName='"&BigClass&"' order by AddDate desc"
rs.Open sql,conn,1,1
end if
if rs.eof and rs.bof then
  response.Write("暂时没有记录")
else 
%>
                  <% 
rs.PageSize=20
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.AbsolutePage=page  
for j=1 to rs.PageSize 
%>
                  <tr>
                    <td width="6%" height="30" align="center" style="BORDER-bottom: #E7E7E7 1px dashed"><img src="Img/arrow_6.gif" width="11" height="11"></td>
                    <td width="65%" height="24" style="BORDER-bottom: #E7E7E7 1px dashed"><% if rs("FirstImageName")<>"" then response.write "<img src='images/news.gif' border=0 alt='图片新闻'>" end if %>
                        <a href="shownews.asp?id=<%= RS("id") %>" target="_blank"><%= RS("TITLE") %></a>　</td>
                    <td width="29%" style="BORDER-bottom: #E7E7E7 1px dashed"><font color="#999999">[<%=FormatDateTime(RS("AddDate"),2)%>] (点击<font color="#ff0000"><%= RS("hits") %></font>) </font></td>
                  </tr>
                  <%
rs.movenext
if rs.eof then exit for
next
%>
                  <tr valign="bottom">
                    <td height="50" colspan="3" align="center" ><form method=Post action="NewsClass.asp?BigClass=<%=BigClass%>&SmallClass=<%=SmallClass%>">
                        <%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=NewsClass.asp?BigClass="&BigClass&"&SmallClass="&SmallClass&"&page=1>首页</a>&nbsp;"
    response.write "<a href=NewsClass.asp?BigClass="&BigClass&"&SmallClass="&SmallClass&"&page=" & Page-1 & ">上一页</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=NewsClass.asp?BigClass="&BigClass&"&SmallClass="&SmallClass&"&page=" & (page+1) & ">"
    response.write "下一页</a> <a href=NewsClass.asp?BigClass="&BigClass&"&SmallClass="&SmallClass&"&SmallClass="&SmallClass&"&page="&rs.pagecount&">尾页</a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#FF0000'>"&rs.recordcount&"</font></b>条记录 <b>"&rs.pagesize&"</b>条记录/页"
   response.write " 转到：<input type='text' name='page' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value=' Goto '  name='cndok'></span></p>"     
%>
                    </form></td>
                  </tr>
                  <% 
end if
rs.close
set rs=nothing
%>
                </table></DIV>
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


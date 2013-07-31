<!--#include file="Inc/SysProduct.asp" -->


<%
function cutstr(tempstr,tempwid)
if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&"..."
else
cutstr=tempstr
end if
end function%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="Head.asp" -->

<body>




<DIV id=index_container>
<DIV id=index_left>
<DIV id=index_search>
<DIV id=index_search1></DIV>
<DIV id=index_search2>
<DIV class=query><table width="95%" border="0" cellspacing="0" cellpadding="0">
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
</DIV>
</DIV>
<DIV id=index_search3></DIV></DIV>
<DIV id=index_prolist>
<DIV id=index_prolist1></DIV>
<DIV id=index_prolist2>
<DIV id=elem-FrontProductCategory_slideTree-001 name="商品分类"><LINK 
href="imgbyy/css2.css" type=text/css 
rel=stylesheet>

<DIV class=FrontProductCategory_slideTree-default 
id=FrontProductCategory_slideTree-001>
<DIV class=left_bottom>
<DIV class=right_bottom>
<DIV class=left_top>
<DIV class=right_top>
<table width="95%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="3"></td>
  </tr>
</table>

<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><% call ShowSmallClass_Tree() %>
</td>
  </tr>
</table>
 <table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10"></td>
  </tr>
</table>

  
</DIV></DIV></DIV></DIV></DIV>
</DIV></DIV>
<DIV id=index_prolist3></DIV></DIV>
<DIV id=index_contact>
<DIV id=index_contact1></DIV>
<DIV id=index_contact2>
<DIV id=elem-FrontSpecify_show-001 name="联系我们说明页"><LINK 
href="imgbyy/css3.css" type=text/css rel=stylesheet>
<DIV class=FrontSpecify_show-default id=FrontSpecify_show-001>
<DIV class=content>
<DIV class=style1><FONT face=Arial>地址：<FONT face=Arial><%=Indexadd%><FONT 
face=Arial></FONT><BR></FONT></FONT>电话：<%=Indextel%> 
<BR>传真：<%=Indexfax%><FONT face=Arial></FONT> 
</DIV></DIV></DIV></DIV></DIV></DIV></DIV>
<DIV id=index_right>
<DIV id=index_news1>
<DIV id=index_news1_1>
<DIV class=index_news_title>企业新闻</DIV>
<DIV class=index_more1><A 
href="NewsClass.asp?BigClass=企业新闻"><IMG src="imgbyy/btn_more.gif" border=0></A></DIV>
</DIV>
<DIV id=index_news1_2>
<DIV >
<DIV class=FrontInfo_listByAsync-05 >
<%
set rs_news=server.createobject("adodb.recordset")
sqltext4="select top 6 * from news where BigClassName='企业新闻' order by id desc"
rs_news.open sqltext4,conn,1,1				 	
%><table width="90%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="7"></td>
  </tr>
</table>

                                      <table 
                              width="95%" border=0 align="center" cellpadding=0 cellspacing=0>
                                        <tbody>
                                          <%i=0
do while not rs_news.eof%>
                                          <tr>
                                            <td width="7%" height="25" valign="middle"><img src="imgbyy/dot1.gif" width="11" height="7" /></td>
                                            <td width="23%" valign="middle"><span style="color: #87888A"><%=FormatDateTime(RS_news("AddDate"),2)%></span></td>
                                            <td width="70%" valign="middle">
								
								<a href="shownews.asp?id=<%=rs_news("id")%>" class="linknews1"><%=cutstr(rs_news("title"),14)%></a>
								
								</td>
                                          </tr>
                                          <%rs_news.movenext
i=i+1
if i=6 then exit do
loop
rs_news.close %>
                                        </tbody>
            </table>
</DIV></DIV></DIV>
<DIV id=index_news1_3></DIV></DIV>
<DIV id=index_news2>
<DIV id=index_news2_1>
<DIV class=index_news_title>业内资讯</DIV>
<DIV class=index_more1><A 
href="NewsClass.asp?BigClass=业内资"><IMG src="imgbyy/btn_more.gif" border=0></A></DIV>
</DIV>
<DIV id=index_news2_2>
<DIV id=elem-FrontInfo_listByAsync-气体知识 name="信息列表"><LINK 
href="imgbyy/css1.css" type=text/css rel=stylesheet>
<DIV class=FrontInfo_listByAsync-05>
<%
set rs_news=server.createobject("adodb.recordset")
sqltext4="select top 6 * from news where BigClassName='业内资讯' order by id desc"
rs_news.open sqltext4,conn,1,1				 	
%><table width="90%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="7"></td>
  </tr>
</table>

                                      <table 
                              width="95%" border=0 align="center" cellpadding=0 cellspacing=0>
                                        <tbody>
                                          <%i=0
do while not rs_news.eof%>
                                          <tr>
                                            <td width="7%" height="25" valign="middle"><img src="imgbyy/dot1.gif" width="11" height="7" /></td>
                                            <td width="23%" valign="middle"><span style="color: #87888A"><%=FormatDateTime(RS_news("AddDate"),2)%></span></td>
                                            <td width="70%" valign="middle">
								
								<a href="shownews.asp?id=<%=rs_news("id")%>" class="linknews1"><%=cutstr(rs_news("title"),14)%></a>
								
								</td>
                                          </tr>
                                          <%rs_news.movenext
i=i+1
if i=6 then exit do
loop
rs_news.close %>
                                        </tbody>
            </table>

</DIV></DIV></DIV>
<DIV id=index_news2_3></DIV></DIV>
<DIV id=index_hotproduct>
<DIV id=index_hotproduct_1></DIV>
<DIV id=index_hotproduct_2><span id="tb_A101">
              <!--------------------显示最新产品的子程序开始--------------------->
                            </span>
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="10"></td>
    </tr>
  </table>
  <span>
        <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
          <%
set rs_Product=server.createobject("adodb.recordset")
sqltext="select top 10 * from Product where Passed=True order by UpdateTime desc"
rs_Product.open sqltext,conn,1,1
%>
          <tr align="center">
            <%
           If rs_Product.eof and rs_Product.bof then
           response.write "<td><p align='center'><font color='#ff0000'>还没任何产品</font></p></td>"
           else
			  row_count=2
			  Do While Not rs_Product.EOF%>
            <td><table width="1%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><table width="100%"  border="0" cellpadding="3" cellspacing="1" bgcolor="#E9E9E9">
                        <tr>
                          <td width="130" height="106" bgcolor="#FFFFFF"><table width="100"  border="0" cellpadding="0" cellspacing="0" bgcolor="#E9E9E9">
                              <tr>
                                <td width="100" height="100" align="center" valign="middle" bgcolor="#FFFFFF"><a href="ProductShow.asp?ID=<%=rs_Product("id")%>"><img src=<%=rs_Product("DefaultPicUrl")%> width="111" height="100" border=0 ></a></td>
                              </tr>
                          </table></td>
                        </tr>
                  </table></td>
                  <td valign="bottom"><img src="Img/yyri2.gif" width="7" height="112" /></td>
                </tr>
                <tr>
                  <td colspan="2"><table width="100%"  border="0" cellspacing="2" cellpadding="0">
                        <tr>
                          <td height="20" align="center"><a href="Product_Show.asp?ID=<%=rs_Product("id")%>"><%=rs_Product("Title")%></a></td>
                        </tr>
                  </table></td>
                </tr>
            </table></td>
            <% if row_count mod 6 =0 then%>
          </tr>
          <tr>
            <%end if%>
            <%
rs_Product.MoveNext
row_count=row_count+1
Loop
end if
rs_Product.close
%>
          </tr>
        </table>
        <!--------------------显示最新产品的子程序结束--------------------->
        </span></DIV>
</DIV>
<DIV class=clear></DIV>
<DIV id=index_hotproduct_3></DIV></DIV></DIV>
<DIV class=clear></DIV>

<DIV id=footer>
<DIV id=elem-FrontSpecify_show-020 ><LINK 
href="imgbyy/css3.css" type=text/css rel=stylesheet>
<DIV class=FrontSpecify_show-default id=FrontSpecify_show-020>
<DIV class=content>
<!--#include file="Inc/Foot.asp" -->


</DIV></DIV></DIV></DIV>

</BODY></HTML>
<!--#include file="Inc/SysProduct.asp" -->
<%
set rsjob=server.createobject("adodb.recordset")
sql="select * from HrDemand order by id desc"
rsjob.open sql,conn,1,1

dim MaPerPage
MaPerPage=4

dim text,checkpage
text="0123456789"
rsjob.PageSize=MaPerPage
for i=1 to len(request("page"))
checkpage=instr(1,text,mid(request("page"),i,1))
if checkpage=0 then
exit for
end if
next

If checkpage<>0 then
If NOT IsEmpty(request("page")) Then
CurrentPage=Cint(request("page"))
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > rsjob.PageCount Then CurrentPage = rsjob.PageCount
Else
CurrentPage= 1
End If
If not rsjob.eof Then rsjob.AbsolutePage = CurrentPage end if
Else
CurrentPage=1
End if
%>




<!-- #include file="Head.asp"-->
<body>

<DIV id=index_container>
  <div id="index_left">
    <div id="left">
      <div id="left_title">企业简介</div>
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
sql="select Title from Aboutus where Language='1' order by Aboutusorder"
rslist.open sql,conn,1,3
do while not rslist.eof
%>
  <LI>
  <a href="Aboutus.asp?Title=<%=rslist("Title")%>"><%=rslist("Title")%></a>
  
  </li>
  
  
  
  <%rslist.movenext 
loop%>


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
  <DIV id=right_title>招聘信息</DIV>
<DIV class=text><table width="98%" height="1%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="206" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
if not rsjob.eof then
i=0
do while not rsjob.eof
%>
        <tr>
          <td 
                  height=1><br>
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#EEEEEE" >
                <tr bgcolor="#DFDFDF">
                  <td width="174" height="30" align="center" bgcolor="#F7F7F7" >职位名称：</td>
                  <td colspan="2" bgcolor="#FFFFFF" >&nbsp;<%=rsjob("HrName")%> </td>
                  <td width="238" bgcolor="#FFFFFF" >&nbsp;<a href="HrDemandAccept.asp?Quarters=<%=rsjob("HrName")%>"><font color="#FF0000">应聘此岗位</font></a> </td>
                </tr>
                <tr bgcolor="#DFDFDF">
                  <td width="174" height="30" align="center" bgcolor="#F7F7F7" >工作地点：</td>
                  <td colspan="3" valign="top" bgcolor="#FFFFFF" >&nbsp;<%=rsjob("HrAddress")%> </td>
                </tr>
                <tr bgcolor="#DFDFDF">
                  <td width="174" height="30" align="center" bgcolor="#F7F7F7" >工资待遇：</td>
                  <td width="198" bgcolor="#FFFFFF">&nbsp;<%=rsjob("HrSalary")%></td>
                  <td width="101" align="center" bgcolor="#FFFFFF" >发布日期：</td>
                  <td bgcolor="#FFFFFF" >&nbsp;<%=rsjob("HrDate")%> </td>
                </tr>
                <tr bgcolor="#DFDFDF">
                  <td height="30" align="center" bgcolor="#F7F7F7" >需求人数：</td>
                  <td align="center" bgcolor="#FFFFFF"><div align="left">&nbsp;<%=rsjob("HrRequireNum")%> 人</div></td>
                  <td align="center" bgcolor="#FFFFFF">有效期限：</td>
                  <td align="center" bgcolor="#FFFFFF"><div align="left">&nbsp;<%=rsjob("HrValidDate")%></div></td>
                </tr>
                <tr bgcolor="#DFDFDF">
                  <td width="174" height="30" align="center" bgcolor="#F7F7F7" ><font color=#000066>具体要求：</font></td>
                  <td colspan="3" align="center" bgcolor="#FFFFFF"><table width="100%"  border="0" cellpadding="5" cellspacing="0" >
                      <tr>
                        <td class="News-05"><%=rsjob("HrDetail")%></td>
                      </tr>
                  </table></td>
                </tr>
                <tr bgcolor="#F7F7F7">
                  <td height="10" colspan="4"><img src="Images/Hrbackup.gif" width="96" height="10"></td>
                </tr>
            </table></td>
        </tr>
        <% 
i=i+1 
if i >= MaPerpage then exit do 
rsjob.movenext 
loop 
end if 
%>
    </table></td>
  </tr>
  <tr>
    <td 
                  height=25 align="center"><%
Response.write "全部"
Response.write "共" & Cstr(rsjob.RecordCount) & "条招聘信息&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "第" & Cstr(CurrentPage) &  "/" & Cstr(rsjob.pagecount) & "&nbsp;"
If currentpage > 1 Then
response.write "<a href='HrDemand.asp?&page="+cstr(1)+"'>&nbsp;首页&nbsp;</a>"
Response.write "<a href='HrDemand.asp?page="+Cstr(currentpage-1)+"'>&nbsp;上一页&nbsp;</a>"
Else
Response.write "&nbsp;上一页&nbsp;"
End if
If currentpage < rsjob.PageCount Then
Response.write "<a href='HrDemand.asp?page="+Cstr(currentPage+1)+"'>&nbsp;下一页&nbsp;</a>"
Response.write "<a href='HrDemand.asp?page="+Cstr(rsjob.PageCount)+"'>&nbsp;尾页&nbsp;</a>"
Else
Response.write ""
Response.write "&nbsp;下一页&nbsp;"
End if
Response.write "转到第"
response.write"<select name='sel_page' onChange='javascript:location=this.options[this.selectedIndex].value;'>"
    for i = 1 to rsjob.PageCount
       if i = currentpage then 
          response.write"<option value='HrDemand.asp?page="&i&"&id="&id&"' selected>"&i&"</option>"
       else
          response.write"<option value='HrDemand.asp?page="&i&"&id="&id&"'>"&i&"</option>"
       end if
    next
response.write"</select>页"
rsjob.close
%>
    </td>
  </tr>
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

<%
rsjob.close
set rsjob=nothing
%>

</BODY></HTML>

<!--#include file="Inc/SysProduct.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="Head.asp" -->

<%
UserName=trim(request("UserName"))
if session("UserName")="" then
	response.redirect "UserServer.asp"
end if
%>
<body>

<DIV id=index_container>
  <div id="index_left">
    <div id="left">
      <div id="left_title">�ͻ�����</div>
      <div id="left_content">
        <DIV id=elem-FrontProductCategory_slideTree-001 name="��Ʒ����"><LINK 
href="imgbyy/css2.css" type=text/css 
rel=stylesheet>

<DIV class=FrontProductCategory_slideTree-default 
id=FrontProductCategory_slideTree-001>
<DIV class=left_bottom>
<DIV class=right_bottom>
<DIV class=left_top>
<DIV class=right_top>


<UL id=treemenu1>


  <LI>
  <a href="UserEdit.asp">�޸Ļ�Ա����</a>
    </li>
  
    <LI>
  <a href="UserEditPwd.asp">�޸Ļ�Ա����</a>
  </li>
     <LI>
  <a href="UserOrder.asp">���ﶩ����ѯ</a>
  </li>
    <LI>
  <a href="FeedbackMember.asp">վ����������</a>
  </li>
    <LI>
  <a href="UserLogout.asp?Action=Ch">�˳���Ա����</a>
  </li>
 
  


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

    <div id="index_contact">
      <div id="index_contact1"></div>
      <div id="index_contact2">
        <div id="elem-FrontSpecify_show-001" name="��ϵ����˵��ҳ">
          <link 
href="about_files/css3.css" type="text/css" rel="stylesheet" />
          <div class="FrontSpecify_show-default" id="FrontSpecify_show-001">
            <div class="content">
              <div class="style1"><font face="Arial">��ַ��<font face="Arial"><%=Indexadd%><font 
face="Arial"></font><br />
                </font></font>�绰��<%=Indextel%> <br />
                ���棺<%=Indexfax%><font face="Arial"></font> </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
  <DIV id=index_right><!--<div id="right_title">��˾���</div> -->
<DIV id=right_content>
<DIV ><LINK 
href="about_files/FrontCommonContent_showDetailByList.css" type=text/css 
rel=stylesheet>
<DIV class=FrontCommonContent_showDetailByList-default 
id=>
<DIV class=content>
  <DIV id=right_title>�û�����</DIV>
<DIV class=text>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <TR>
                          <TD align="center"><%
UserName=Session("UserName")
set rs=server.createobject("adodb.recordset")
sql="select * from Feedback where UserName='"&UserName&"' order by id desc"
rs.open sql,conn,1,1

dim MaxPerPageFeedback
MaxPerPageFeedback=5

'ȡ��ҳ��,���ж�����������Ƿ��������͵����ݣ��粻�ǽ��Ե�һҳ��ʾ
text="0123456789"
Rs.PageSize=MaxPerPageFeedback
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
If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount
Else
CurrentPage= 1
End If
If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
CurrentPage=1
End if

call list

'��ʾ���ӵ��ӳ���
Sub list()%>
                              <%
If rs.eof and rs.bof then
  response.write "<br><br><p algin=center>�Բ�������û�и���������</p><br><a href='Feedback.asp'>�����������</a>"
  response.end
End if
%>
                              <br>
                              <%
i=0
do while not rs.eof
%>
                              <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF" >
                                <tr bgcolor="#EBEBEB">
                                  <td width="17%" height="40" align="right" bgcolor="#F5F5F5" > �����⣺ </td>
                                  <td colspan="5" bgcolor="#F5F5F5">&nbsp;
                                      <%if rs("username")="����Ա"then%>
                  [����Ա����]
                  <%end if%>
                  <%=rs("title")%></td>
                                </tr>
                                <tr bgcolor="#DFDFDF">
                                  <td height="22" align="right" bgcolor="#F5F5F5" > �������ݣ� </td>
                                  <td colspan="5" align="center" bgcolor="#F5F5F5" ><table width="97%"  border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                        <td height="40" ><%=rs("Content")%> </td>
                                      </tr>
                                      
                                  </table></td>
                                </tr>
                                <tr bgcolor="#EFEFEF">
                                  <td height="22" align="right" bgcolor="#F5F5F5" > �����ߣ� </td>
                                  <td width="22%" bgcolor="#F5F5F5" ><%=rs("Receiver")%> </td>
                                  <td width="12%" align="right" bgcolor="#F5F5F5" >����ʱ�䣺</td>
                                  <td width="17%" bgcolor="#F5F5F5" ><%=rs("Time")%></td>
                                  <td width="12%" align="right" bgcolor="#F5F5F5" >�ظ�ʱ�䣺</td>
                                  <td width="20%" bgcolor="#F5F5F5" ><%=rs("Retime")%> </td>
                                </tr>
                                <tr bgcolor="#DFDFDF">
                                  <td height="22" align="right" bgcolor="#F5F5F5" > ����Ա�ظ��� </td>
                                  <td colspan="5" align="center" bgcolor="#F5F5F5" ><table width="97%"  border="0" cellpadding="0" cellspacing="0" >
                                      <tr>
                                        <td height="4"></td>
                                      </tr>
                                      <tr>
                                        <td height="40" ><%if rs("ReFeedback")=""then%>
                        ["δ�ظ�"]
                        <%else%>
                        <%=rs("ReFeedback")%>
                        <%end if%>
                                        </td>
                                      </tr>
                                      <tr>
                                        <td height="4"></td>
                                      </tr>
                                  </table></td>
                                </tr>
                                <tr valign="top" bgcolor="#F7F7F7">
                                  <td height="12" colspan="6" align="left" ><img src="Images/Hrbackup.gif" width="96" height="10"></td>
                                </tr>
                            </table>
                              <%
i=i+1
if i >= MaxPerPageFeedback then exit do
rs.movenext
loop
%>
                              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td align="center"><div align="center">
                                      <%
Response.write "ȫ��"
Response.write "��" & Cstr(Rs.RecordCount) & "������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "��" & Cstr(CurrentPage) &  "/" & Cstr(rs.pagecount) & "&nbsp;"
If currentpage > 1 Then
response.write "<a href='FeedbackMember.asp?&page="+cstr(1)+"'>&nbsp;��ҳ&nbsp;</a>"
Response.write "<a href='FeedbackMember.asp?page="+Cstr(currentpage-1)+"'>&nbsp;��һҳ&nbsp;</a>"
Else
Response.write "&nbsp;��һҳ&nbsp;"
End if
If currentpage < Rs.PageCount Then
Response.write "<a href='FeedbackMember.asp?page="+Cstr(currentPage+1)+"'>&nbsp;��һҳ&nbsp;</a>"
Response.write "<a href='FeedbackMember.asp?page="+Cstr(Rs.PageCount)+"'>&nbsp;βҳ&nbsp;</a>"
Else
Response.write ""
Response.write "&nbsp;��һҳ&nbsp;"
End if
Response.write "ת����"
response.write"<select name='sel_page' onChange='javascript:location=this.options[this.selectedIndex].value;'>"
    for i = 1 to Rs.PageCount
       if i = currentpage then 
          response.write"<option value='FeedbackMember.asp?page="&i&"&id="&id&"' selected>"&i&"</option>"
       else
          response.write"<option value='FeedbackMember.asp?page="&i&"&id="&id&"'>"&i&"</option>"
       end if
    next
response.write"</select>ҳ"
%>
                                  </div></td>
                                </tr>
                              </table>
                              <%end sub%>
                          </TD>
                        </TR>
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

<%
rs.close
set rs=nothing
%>    
</BODY></HTML>
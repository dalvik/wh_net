<!--#include file="Inc/SysProduct.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="Head.asp" --><%
Subject=request.form("Subject")
UserName=session("UserName")
Receiver=request.form("Receiver")
Sex=request.form("Sex")
Phone=request.form("Phone")
Add=request.form("Add")
Email=request.form("Email")
Notes=request.form("Notes")
CompanyName=request.form("CompanyName")
Fax=request.form("Fax")

ProductList = Session("ProductList")
if productlist<>"" then
  sql="select Title from Product where Product_Id in ("&productlist&") order by Product_Id"
  Set rs_title = conn.Execute( sql )
else
  response.redirect "error.asp?error=007"
  response.end
end if
%>


<DIV id=index_container>
  <div id="index_left">
    <div id="left">
      <div id="left_title">��Ʒϵ��</div>
      <div id="left_content">
        <DIV id=elem-FrontProductCategory_slideTree-001 name="��Ʒ����"><LINK 
href="imgbyy/css2.css" type=text/css 
rel=stylesheet>
          <div class="FrontProductCategory_slideTree-default" 
id="FrontProductCategory_slideTree-001">
            <div class="left_bottom">
              <div class="right_bottom">
                <div class="left_top">
                  <div class="right_top">
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
                  </div>
                </div>
              </div>
            </div>
          </div>
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
  <DIV id=right_title>��Ʒ����</DIV>
<DIV class=text>
  <form action="ordersent.asp" method="post" name="confirm" id="confirm">
    <%=UserName%>:<br />
    ���ã�<br />
    ����������ѯ����Ϣ�����û��������
  <input name="submit" type="submit" value="�����ύ" />
    �����߷���
  <input type="button" name="Submit2" onClick="javascript:history.go(-1)" value="�����޸�" />
  <br />
  &nbsp;&nbsp;���κ������뼰ʱ��������ϵ�� <br />
  <br />
  <input type="hidden" value="ok" name="confirm" />
  <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%" bgcolor="#eaeaea">
    <tr bgcolor="#FFFFFF">
      <td height="40" class="protext">&nbsp;&nbsp;&nbsp;��ϵ��Ϣ</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height="25"><span style="color: #CC0000"> <br />
        &nbsp;&nbsp;&nbsp;&nbsp;��ϵ��:<%=Receiver%><br />
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;�Ա�:
        <% if Sex="1" then response.Write("��") else response.Write("Ů") end if%>
        <br />
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;���⣺<%=Subject%><br />
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;˵����<%=Notes%><br />
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;��˾���ƣ�<%=CompanyName%><br />
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;��˾��ַ��<%=Add%><br />
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;���䣺<%=Email%><br />
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;��˾�绰��<%=Phone%><br />
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;���棺<%=Fax%><br />
        <input type="hidden" name="Receiver" value="<%=Receiver%>" />
        <input type="hidden" name="Subject" value="<%=Subject%>" />
        <input type="hidden" name="Phone" value="<%=Phone%>" />
        <input type="hidden" name="Add" value="<%=Add%>" />
        <input type="hidden" name="Sex" value="<%=Sex%>" />
        <input type="hidden" name="Email" value="<%=Email%>" />
        <input type="hidden" name="Notes" value="<%=Notes%>" />
        <input type="hidden" name="CompanyName" value="<%=CompanyName%>" />
        <input type="hidden" name="Fax" value="<%=Fax%>" />
        <br />
      </span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="50%" height="40" align="center" class="blue" style="font-weight: bold"><font color="#CC0000">�� Ʒ �� ��</font></td>
    </tr>
    <%  
  While Not rs_title.Eof    
%>
    <tr bgcolor="#FFFFFF">
      <td width="50%" height="40" align="center" class="boldblue"><font color="#CC0000"><%=rs_title("Title")%></font></td>
    </tr>
    <%
rs_title.MoveNext
Wend
rs_title.close
set rs_title=nothing
%>
  </table>
  <br />
  </form>
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
<%
set conn=nothing
%>
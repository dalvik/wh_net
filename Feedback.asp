<!--#include file="Inc/SysProduct.asp" -->
<%
UserName=Session("UserName")
set rs = Server.CreateObject("ADODB.recordset")
sql="select * from User where UserName='"&UserName&"'"
rs.open sql,conn,1,1
%>
<!-- #include file="Head.asp" -->

<DIV id=index_container>
  <div id="index_left">
    <div id="left">
      <div id="left_title">��������</div>
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
<a href="FeedbackView.asp">�鿴����</a>    </li>
  
    <LI>
<a href="Feedback.asp">��Ҫ����</a>  </li>
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
  <DIV id=right_title>���Է���</DIV>
<DIV class=text>
  <form method="post" action="FeedbackSave.asp?Language=ch">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>�𾴵Ŀͻ���
          <p>��������������ǵĲ�Ʒ��������κ�����ͽ����뼰ʱ�������ǣ����ǽ������������Ĵ𸴡�<br />
            �����ע��һ����Ա�ţ��Ժ�ÿ������ʱֻҪ��¼�ٲ����ظ���д�����ϵ��Ϣ�ˣ�����ͨ����Ա����ɼ��в鿴ֻ�����������ԡ�����ע���ԱҲ��ֱ�����ԣ� <br />
            <br />
          </p></td>
      </tr>
    </table>
    <table width="100%" height="409"
border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E3E3E3">
      <tr>
        <input type="hidden" name="Username" value="<%=Username%>" />
        <td height="25" align="right" bgcolor="#FFFFFF">���⣺ </td>
        <td height="25" bgcolor="#FFFFFF"><input type="text" name="Title" size="42" maxlength="36" style="font-size: 14px">
          *</td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">���� *��</td>
        <td height="25" bgcolor="#FFFFFF"><textarea rows="10" name="Content" cols="45" style="font-size: 14px" ></textarea></td>
      </tr>
      <tr>
        <td width="23%" height="25" align="right" bgcolor="#FFFFFF">��˾���ƣ�</td>
        <td width="77%" height="-6" bgcolor="#FFFFFF"><font>
          <input type="text" name="CompanyName" size="30" maxlength="36" value="<%=rs("CompanyName")%>" style="font-size: 14px">
        </font>* </td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">��˾��ַ��</td>
        <td height="-2" bgcolor="#FFFFFF"><font>
          <input name="Add" type="text"  id="Add" style="font-size: 14px" value="<%=rs("Add")%>" size="40" maxlength="60" />
        </font></td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">�ʱࣺ</td>
        <td height="-2" bgcolor="#FFFFFF"><font>
          <input name="Postcode" type="text"  id="Postcode" style="font-size: 14px" value="<%=rs("Postcode")%>" size="12" maxlength="6" />
        </font></td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">��ϵ�ˣ�</td>
        <td width="77%" height="-2" bgcolor="#FFFFFF"><font>
          <input type="text" name="Receiver" size="12" maxlength="30" value="<%=rs("Receiver")%>" style="font-size: 14px">
        </font> * </td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">��ϵ�绰��</td>
        <td width="77%" height="-1" bgcolor="#FFFFFF"><font>
          <input type="text" name="Phone" size="24" maxlength="36" value="<%=rs("Phone")%>" style="font-size: 14px">
        </font>* </td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">�ֻ���</td>
        <td height="11" bgcolor="#FFFFFF"><font>
          <input name="Mobile" type="text" id="Mobile" style="font-size: 14px" value="<%=rs("Mobile")%>" size="24" maxlength="36" />
        </font></td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">��ϵ���棺</td>
        <td width="77%" height="11" bgcolor="#FFFFFF"><font>
          <input type="text" name="Fax" size="18" maxlength="36" value="<%=rs("Fax")%>" style="font-size: 14px">
        </font></td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">E-mail��</td>
        <td height="11" bgcolor="#FFFFFF"><font>
          <input type="text" name="Email" size="18" maxlength="36" value="<%=rs("Email")%>" style="font-size: 14px">
        </font></td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">���Ļ���</td>
        <td width="77%" height="11" bgcolor="#FFFFFF"><% if session("username")<>"" then  %>
            <input type="radio" name="Publish" value="1" />
          ��
          <input name="Publish" type="radio" value="0" checked="checked" />
          ��
          <% else %>
          <input name="Publish" type="hidden" value="0" checked="checked" />
          <%	  response.write "<font color='#ff6600'>��ע���Ա������</font>"
					end if%>
          (ֻ�й���Ա���Լ��ܿ�)</td>
      </tr>
      <tr>
        <td width="23%" height="0" valign="top" bgcolor="#FFFFFF"></td>
        <td width="77%" height="0" valign="top" bgcolor="#FFFFFF"><input type="submit" value="�ύ����"
name="cmdOk" />
            <input type="reset" value="��д" name="cmdReset" />
        </td>
      </tr>
    </table>
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
<%
rs.close
set rs=nothing
%>      
</BODY></HTML>

<%@ LANGUAGE="VBScript" %>
<!--#include file="Inc/Conn.asp" -->
<!--#include file="Inc/Function.asp"-->
<%
OrderNum= Request("OrderNum")
IF Session("UserName")="" Then
 response.redirect "userserver.asp"
Else
set Rs3 = Server.CreateObject("ADODB.recordset")
sql3="select * from OrderList where OrderNum='"&OrderNum&"'"
rs3.open sql3,conn,1,1
IF  rs3.RecordCount >=1  then
IF Session("UserName")=rs3("UserName") Then
%>
<html>
<head>
<title>�ͻ�ѯ����ϸ��Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="mt_style.css" rel="stylesheet" type="text/css">
</head>
<%
id=Form_Id
set rs=server.createobject("adodb.recordset")
sqltext="select * from OrderList where OrderNum='"&OrderNum&"'"
rs.open sqltext,conn,1,1
%>

<body>
<br>    <div align="center">
<center>

  <TABLE cellSpacing=1 cellPadding=4 width=530 bgColor=#eaeaea>
    <TBODY>
      <TR bgColor=#eeeeee> 
        <TD height="32"  colSpan=2><div align="center"><font color="#000000"><strong>�ͻ�ѯ����ϸ����</strong></font></div></TD>
      </TR>
      <TR> 
        <TD  width=127 bgColor=#DBDBDB height=12 align="right"><font color="#000000">ѯ�ۺţ�</font></TD>
        <TD  width=402 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("OrderNum")%></TD>
      </TR>
      <TR> 
        <TD bgColor=#DBDBDB height=12 align="right">��˾����<font color="#000000">��</font></TD>
        <TD  width=402 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("CompanyName")%></TD>
      </TR>
      <tr> 
        <TD  width=127 bgColor=#DBDBDB height=25 align="right"><font color="#000000">��ϵ�ˣ�</font></TD>
        <TD  width=402 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Receiver")%></TD>
      </TR>
      <tr> 
        <TD  width=127 bgColor=#DBDBDB height=25 align="right"><font color="#000000">��˾��ַ��</font></TD>
        <TD  width=402 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Add")%></TD>
      </tr>
      <tr> 
        <TD  width=127 bgColor=#DBDBDB height=12 align="right"><font color="#000000">��ϵ�绰��</font></TD>
        <TD  width=402 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("Phone")%></TD>
      </tr>
      <tr> 
        <TD bgColor=#DBDBDB height=12 align="right">��ϵ����<font color="#000000">��</font></TD>
        <TD  width=402 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("Fax")%></TD>
      </tr>
      <tr> 
        <TD  width=127 bgColor=#DBDBDB height=25 align="right"><font color="#000000">�������䣺</font></TD>
        <TD  width=402 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("EMail")%></TD>
      </tr>
      <tr> 
        <TD  width=127 height=25 align="right" bgColor=#DBDBDB><font color="#000000">��ע��</font></TD>
        <TD  width=402 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Notes")%></TD>
      </tr>
      <tr> 
        <TD  width=127 bgColor=#DBDBDB height=24 align="right"><font color="#000000">ѯ�����ڣ�</font></TD>
        <TD  width=402 height=24 bgcolor="#eeeeee">&nbsp; <%=rs("OrderTime")%></TD>
      </tr>
      <tr> 
        <TD  width=127 bgColor=#DBDBDB height=25 align="right"><font color="#000000">ѯ���Ƿ��Ѿ�������</font></TD>
        <TD  width=402 height=25 bgcolor="#eeeeee">&nbsp;  <%If rs("Flag")="Yes" Then%>
                                          �Ѵ���
                                          <%else%>
                                          <font color="#FF0000"> δ����</font> 
                                          <%End If%> </TD>
      </tr>
      <TR> 
        <TD height="31"  colSpan=2 bgcolor="#eeeeee"> <p align="center"><font color="#000000">��Ʒ</font>�б�</p></TD>
      </TR>
      <%
set rs2=server.createobject("adodb.recordset")
sqltext2 = "select A.Product_Id,A.OrderNum,A.ProductUnit,A.BuyPrice,C.Title,C.Price,C.BigClassName,C.SmallClassName from OrderDetail A,Product C where A.OrderNum='"&OrderNum&"' and A.Product_Id=C.Product_Id"

'sqltext2="select * from OrderDetail where OrderNum='"&OrderNum&"'"
rs2.open sqltext2,conn,1,1
%>
      <TR> 
        <TD height="15"  colSpan=2 valign="top" bgcolor="#eeeeee"> <div align="center"> 
            <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#cccccc" bordercolordark="#eeeeee"  height="66">
              <tr> 
                <td width="18%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">��Ʒ���</font></td>
                <td width="23%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">��Ʒ����</font></td>
                <td width="24%" bgcolor="#DBDBDB" align="center"><font color="#000000">��Ʒ</font>�۸�</td>
                <td width="35%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">��Ʒ���</font></td>
              </tr>
              <%			 
While Not rs2.EOF     
	   '�����ܽ��%>
              <tr> 
                <td width="18%" align="center" height="22"><%=rs2("Product_Id")%></td>
                <td width="23%" align="center" height="22"><%=rs2("Title")%></td>
                <td width="24%" align="center"><%=rs2("BuyPrice")%></td>
                <td width="35%" align="center" height="22"><%=rs2("BigClassName")%> => <%=rs2("SmallClassName")%></td>
              </tr>
              <%
rs2.MoveNext
Wend
%>
            </table>
          </div></TD>
      </TR><center>
    <TR> 
      <TD height="25"  colSpan=2 bgcolor="#eeeeee"> <p align="center"> 
          <input name="close" type="button" id="close" onClick="javascript:window.close()" value="�ر�">
      </TD>
    </TR>
  </TABLE>
</div>
</form>

<p>
<%
Else
errmsg=errmsg+"<li>�����ܲ鿴����������ѯ�ۺţ��������������Լ���ѯ�ۺţ�</li>"
call WriteErrMsg()
End If
Else
errmsg=errmsg+"<li>�������ѯ�ۺŲ����ڻ��ʽ����ȷ�����������룡</li>"
call WriteErrMsg()
End IF
End if
rs3.close
conn.close
%>

</body>
</html>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
OrderNum=request("OrderNum")
if Request.QueryString("action")="southidc" then
 id=request("id")
 price=request.form("price")
 Set rs_save = Server.CreateObject("ADODB.Recordset")
 rs_save.open "select * from OrderDetail where ID="&id,conn,1,3
 rs_save("BuyPrice")=price
 rs_save.update
 rs_save.close
 set rs_flag=server.createobject("adodb.recordset")
 sqltext="Update OrderList set Flag='Yes' where OrderNum='"&OrderNum&"'"
 rs_flag.open sqltext,conn,1,3 
 rs_flag.close
 'response.redirect "Admin_OrderDetail.asp?OrderNum='"&OrderNum&"'"
end if 

page=request("page")
set rs=server.createobject("adodb.recordset")
sqltext="select * from OrderList where OrderNum='"&OrderNum&"'"
rs.open sqltext,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <strong><br>
      </strong> <table  width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" height="25"> 
            <div align="center"><strong><%=rs("OrderNum")%>�Ŵ������� <br>
              </strong></div></td>
        </tr>
        <tr>          
            <td> <TABLE cellSpacing=1 cellPadding=4 width=562 bgColor=#006699 height="436">
                <TBODY>
                  <TR> 
                    <TD height="10"  colSpan=2 valign="top" bgcolor="#DBDBDB"></TD>
                  </TR>
                  <TR bgColor=#eeeeee> 
                    <TD height="32"  colSpan=2><font color="#000000">�ͻ�ѯ�۵���ϸ����</font></TD>
                  </TR>
                  <TR> 
                    <TD  width=108 bgColor=#DBDBDB height=25 align="right"><font color="#000000">ѯ�۵��ţ�</font></TD>
                    <TD  width=433 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("OrderNum")%></TD>
                  </TR>
                  <tr> 
                    <TD  width=108 bgColor=#DBDBDB height=25 align="right">��˾����<font color="#000000">��</font></TD>
                    <TD  width=433 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("CompanyName")%></TD>
                  </TR>
                  <tr> 
                    <TD  width=108 bgColor=#DBDBDB height=25 align="right"><font color="#000000">��ϵ�ˣ�</font></TD>
                    <TD  width=433 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Receiver")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=108 bgColor=#DBDBDB height=25 align="right"><font color="#000000">�ջ���ַ��</font></TD>
                    <TD  width=433 height=-2 bgcolor="#eeeeee">&nbsp; <%=rs("Add")%></TD>
                  </tr>
                  <tr> 
                    <TD bgColor=#DBDBDB height=25 align="right"><font color="#000000">��ϵ�绰��</font></TD>
                    <TD  width=433 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Phone")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=108 bgColor=#DBDBDB height=25 align="right">��ϵ����<font color="#000000">��</font></TD>
                    <TD  width=433 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Fax")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=108 bgColor=#DBDBDB height=25 align="right"><font color="#000000">�������䣺</font></TD>
                    <TD  width=433 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("EMail")%></TD>
                  </tr>
                  <tr> 
                    <TD height=25 align="right" bgColor=#DBDBDB>����<font color="#000000">��</font></TD>
                    <TD height=25 bgcolor="#eeeeee"> &nbsp; <%=rs("Subject")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=108 height=25 align="right" bgColor=#DBDBDB><font color="#000000">��ע��</font></TD>
                    <TD  width=433 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Notes")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=108 bgColor=#DBDBDB height=24 align="right"><font color="#000000">ѯ�����ڣ�</font></TD>
                    <TD  width=433 height=24 bgcolor="#eeeeee">&nbsp; <%=rs("OrderTime")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=108 bgColor=#DBDBDB height=25 align="right"><font color="#000000">ѯ���Ƿ��Ѿ�������</font></TD>
                    
                  <TD  width=433 height=25 bgcolor="#eeeeee"> 
                    <%If rs("Flag")="Yes" Then%>
                    �Ѵ��� 
                    <%else%>
                      δ���� <%End If%> </TD>
                  </tr>
                  <TR> 
                    <TD height="31"  colSpan=2 valign="top" bgcolor="#eeeeee"> 
                      <p align="center">ѯ����ƷϸĿ</p></TD>
                  </TR>
                  <%
set rs2=server.createobject("adodb.recordset")
sqltext2 = "select A.OrderNum,A.id,A.Product_Id,A.BuyPrice,A.ProductUnit,"&_		
		"C.Title,C.Price "&_
		" from OrderDetail A,Product C"&_
		" where A.OrderNum='"&OrderNum&"' and A.Product_Id=C.Product_Id"
rs2.open sqltext2,conn,1,1
%>
                  <TR> 
                    <TD height="15"  colSpan=2 valign="top" bgcolor="#eeeeee"> 
                      <div align="center">
					   
                        <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#006699" bordercolordark="#eeeeee"  height="67">
                          <tr> 
                            <td width="38%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">��Ʒ����</font></td>
                            <td width="31%" bgcolor="#DBDBDB" align="center"><font color="#000000">��Ʒ�۸�</font></td>
                            <td width="31%" bgcolor="#DBDBDB" align="center">&nbsp;</td>
                          </tr>
                          <% 						 
While Not rs2.EOF
     
	  '�����ܽ��%>
	  
	   <form name="form" method="post" action="Admin_OrderDetail.asp?action=southidc&id=<%=rs2("id")%>">	    
                          <tr> 						  
                            <td height="22" align="center"><%=rs2("Title")%></td>
                            <td height="22" align="center"><input name="Price" type="text" id="Price" size="10" value=<%=rs2("BuyPrice")%> onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))">
                              Ԫ</td>
                            <td height="22" align="center"><input type="submit" name="Submit" value="�� ��"></td>
                          </tr>
					<INPUT TYPE="hidden" name="OrderNum" value="<%=OrderNum%>">
						  
</form>
                          <%
rs2.MoveNext
Wend
%>
                        </table>
                      </div></TD>
                  </TR>
                <center>
                  <TR> 
                    <TD height="25"  colSpan=2 bgcolor="#eeeeee"> <p align="center"> 
                       
<%
rs.close
rs2.close
%>
                        <input type="button" value="����" name="B4" onclick="javascript:window.history.go(-1)">
                    </TD>
                  </TR>
                  <TR bgColor=#DBDBDB> 
                    <TD height="3"  colSpan=2></TD>
                  </TR>
                </center>
              </TABLE></td>         
        </tr>
      </table></td>
  </tr>
</table>
<script>
	function regInput(obj, reg, inputStr)
	{
		var docSel	= document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")	return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange	= obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
	}
</script>
<!-- #include file="Inc/Foot.asp" -->
<!--#include file="Inc/SysProduct.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="Head.asp" -->
<%
public ProductList
if request("order")="��Ʒ����" then
   Session("ProductList")=""
end if
ProductList = Session("ProductList")
Sub PutToShopBag(Prodid, ProductList)
   If Len(ProductList) = 0 Then
      ProductList = "'" & Trim(Prodid) & "'"
   ElseIf InStr( ProductList, Prodid ) <= 0 Then
      ProductList = ProductList&", '"&Trim(Prodid)&"'"
   End If   
End Sub
%>
<%
Product_Id = Trim(Request("Product_Id"))
if instr(Product_Id,",")>0 then		
		idArr=Split(Product_Id, ",")
		for i = 0 to ubound(idArr)
			PutToShopBag idArr(I), ProductList
		next
else
		   PutToShopBag Product_Id, ProductList
end if
%>

<%
Session("ProductList") = ProductList

'�жϹ��ﳵ�Ƿ�Ϊ��
if Productlist<>"''" then 
  sql="select * from Product where Product_Id in ("&ProductList&") order by     Product_Id"
  Set rs_price = conn.Execute(sql)
else 
  response.redirect "error.asp?error=007"
  response.end
end if
%> 





<script language="JavaScript">
function CheckUserForm()
{
	if (document.payment.Subject.value.length == 0) {
		alert("���ⲻ��Ϊ��.");
		document.payment.Subject.focus();
		return false;
	}	
	if (document.payment.Notes.value.length == 0) {
		alert("˵������Ϊ��.");
		document.payment.Notes.focus();
		return false;
	}
	if (document.payment.Receiver.value.length == 0) {
		alert("��ϵ�˲���Ϊ��.");
		document.payment.Receiver.focus();
		return false;
	}
    if (document.payment.CompanyName.value.length == 0) {
		alert("��˾������Ϊ��.");
		document.payment.CompanyName.focus();
		return false;
	}   	
	if (document.payment.Email.value.length == 0) {
		alert("Email����Ϊ��.");
		document.payment.Email.focus();
		return false;
	}
	if (document.payment.Email.value.length > 0 && !document.payment.Email.value.match( /^.+@.+$/ ) ) {
	    alert("EMAIL ��������.");
	    document.payment.Email.focus();
		return false;
	}	
	if (document.payment.Phone.value.length == 0) {
		alert("�绰����Ϊ��.");
		document.payment.Phone.focus();
		return false;
	}  
	return true;
}
</script>

<body>

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
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td><%If Session("UserName")<>"" Then%>
                  <table width="100%" align="center" cellpadding="0" cellspacing="0">
                    <tbody>
                      <tr valign="top" >
                        <td  width="100%"><div align="center">
                            <center>
                            </center>
                        </div>
                          ��Ա:<%=Session("UserName")%>������ </td>
                      </tr>
                      <tr valign="top" >
                        <td  width="100%">&nbsp;&nbsp;</td>
                      </tr>
                    </tbody>
                  </table>
                <%end if%>
              </td>
            </tr>
          </table>
        <form action="Payment.asp" method="post" name="check" id="check">
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
              <tr bgcolor="#FFFFFF">
                <td colspan="2" height="25">&nbsp;&nbsp;<font color="B0266D"><%=session("UserName")%> </font>��ѡȡ�Ĳ�Ʒ��ϸ����:<br />
                    <font color="B0266D">&nbsp; </font> </td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td width="18%" height="25" align="center">ѡȡ</td>
                <td width="82%" height="25" align="center">��Ʒ����</td>
              </tr>
              <%
 While Not rs_price.EOF   
%>
              <tr bgcolor="#FFFFFF">
                <td width="18%" height="20"><div align="center">
                    <input type="checkbox" name="Product_Id" value="<%=rs_price("Product_Id")%>" checked="checked" />
                </div></td>
                <td width="82%" bgcolor="#FFFFFF"><div align="center"><font color="B0266D"><%=rs_price("Title")%></font></div></td>
              </tr>
              <%
rs_price.MoveNext
Wend
rs_price.close
set rs_price=nothing
%>
              <tr bgcolor="#FFFFFF">
                <td height="25" align="right"><div align="center">
                    <input type="submit" name="order" value="��Ʒ����" />
                </div></td>
                <td height="25" align="right"><div align="center"> </div></td>
              </tr>
            </table>
        </form>
        <%
sql="select * from [User] where UserName='"&session("UserName")&"'"
set rs_price=server.createobject("adodb.recordset")
rs_price.open sql,conn,1,1
Sex=rs_price("Sex")
Phone=rs_price("Phone")
Email=rs_price("Email")
CompanyName=rs_price("CompanyName")
Receiver=rs_price("Receiver")
Add=rs_price("Add")
Fax=rs_price("Fax")
rs_price.close
set rs_price=nothing
%>
          <form action="Orderpreview.asp" method="post" name="payment" id="payment"  onsubmit="return CheckUserForm();">
            <table border="0" cellspacing="1" cellpadding="3" align="center" width="100%" bgcolor="#CCCCCC">
              <tr>
                <td width="96" height="23" bgcolor="#FFFFFF"><div align="right">TO<strong>:</strong></div></td>
                <td width="439" bgcolor="#FFFFFF"><%=CompanyName%>&nbsp;</td>
              </tr>
              <tr>
                <td height="21" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> ����:</div></td>
                <td height="21" bgcolor="#FFFFFF"><input name="Subject" type="text"  size="50" maxlength="100" value="" /></td>
              </tr>
              <tr>
                <td height="21" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> ˵��:<br />
                  (20-4000�ַ�)</div></td>
                <td height="21" bgcolor="#FFFFFF"><textarea name="Notes" cols="60" rows="10"></textarea></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="21" colspan="2"><div align="center">лл������ǵĲ�Ʒ����Ȥ�����ǻᾡ�������ϵ��</div></td>
              </tr>
            </table>
            <br />
            <table border="0" cellspacing="1" cellpadding="3" align="center" width="100%" bgcolor="#CCCCCC">
              <tr bgcolor="#FFFFFF">
                <td width="18%" rowspan="2" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> ��ϵ��<strong>:</strong></div></td>
                <td width="16%"><div align="right">
                    <input name="Sex" type="radio" value="1" <% if Sex="1" then response.Write("checked") end if%> />
                  ����</div></td>
                <td width="66%" colspan="2" rowspan="2"><input name="Receiver" type="text" class="form" id="name" size="30" value="<%=Receiver%>" />
                    <b></b></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td><div align="right">
                    <input type="radio" name="Sex" value="0" <% if Sex="0" then response.Write("checked") end if%> />
                  Ůʿ</div></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="30" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> ��˾����<strong>:</strong></div></td>
                <td height="30" colspan="3" bgcolor="#FFFFFF"><input type="text" name="CompanyName" size="40" value="<%=CompanyName%>" />
                  (4-100 �ַ�) </td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="30" bgcolor="#FFFFFF"><div align="right">��ַ:</div></td>
                <td height="30" colspan="3"><input type="text" name="Add"  size="50" value="<%=Add%>" /></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="30" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> E-mail<strong>:</strong></div></td>
                <td height="30" colspan="3"><input type="text" name="Email"  size="30" value="<%=Email%>" />
                    <b></b></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="30" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> ��˾�绰<strong>:</strong></div></td>
                <td colspan="3"><input type="text" name="Phone" class="form" size="20" value="<%=Phone%>" />
                  (Tel:86-010-81991660)</td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="30" bgcolor="#FFFFFF"><div align="right">����<strong>:</strong></div></td>
                <td colspan="3"><input type="text" name="Fax" size="20" value="<%=Fax%>" />
                  (Format:86-010-81991660)</td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="32" colspan="4" align="center"><input type="submit" name="nextstep" value="�ύ����" />
                  &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
                  <input type="button" value="��������" language="javascript" onClick="javascript:window.close()" name="button" />
                  &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
                  <input name="reset" type="reset" id="reset" value="������д" />
                </td>
              </tr>
            </table>
          </form></td>
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

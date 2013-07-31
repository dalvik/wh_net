<!--#include file="Inc/SysProduct.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="Head.asp" -->
<%
public ProductList
if request("order")="产品更新" then
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

'判断购物车是否为空
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
		alert("标题不能为空.");
		document.payment.Subject.focus();
		return false;
	}	
	if (document.payment.Notes.value.length == 0) {
		alert("说明不能为空.");
		document.payment.Notes.focus();
		return false;
	}
	if (document.payment.Receiver.value.length == 0) {
		alert("联系人不能为空.");
		document.payment.Receiver.focus();
		return false;
	}
    if (document.payment.CompanyName.value.length == 0) {
		alert("公司名不能为空.");
		document.payment.CompanyName.focus();
		return false;
	}   	
	if (document.payment.Email.value.length == 0) {
		alert("Email不能为空.");
		document.payment.Email.focus();
		return false;
	}
	if (document.payment.Email.value.length > 0 && !document.payment.Email.value.match( /^.+@.+$/ ) ) {
	    alert("EMAIL 输入有误.");
	    document.payment.Email.focus();
		return false;
	}	
	if (document.payment.Phone.value.length == 0) {
		alert("电话不能为空.");
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
      <div id="left_title">产品系列</div>
      <div id="left_content">
        <DIV id=elem-FrontProductCategory_slideTree-001 name="商品分类"><LINK 
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
  <DIV id=right_title>产品订购</DIV>
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
                          会员:<%=Session("UserName")%>：　　 </td>
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
                <td colspan="2" height="25">&nbsp;&nbsp;<font color="B0266D"><%=session("UserName")%> </font>您选取的产品明细如下:<br />
                    <font color="B0266D">&nbsp; </font> </td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td width="18%" height="25" align="center">选取</td>
                <td width="82%" height="25" align="center">产品名称</td>
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
                    <input type="submit" name="order" value="产品更新" />
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
                <td height="21" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> 标题:</div></td>
                <td height="21" bgcolor="#FFFFFF"><input name="Subject" type="text"  size="50" maxlength="100" value="" /></td>
              </tr>
              <tr>
                <td height="21" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> 说明:<br />
                  (20-4000字符)</div></td>
                <td height="21" bgcolor="#FFFFFF"><textarea name="Notes" cols="60" rows="10"></textarea></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="21" colspan="2"><div align="center">谢谢你对我们的产品感兴趣，我们会尽快和你联系！</div></td>
              </tr>
            </table>
            <br />
            <table border="0" cellspacing="1" cellpadding="3" align="center" width="100%" bgcolor="#CCCCCC">
              <tr bgcolor="#FFFFFF">
                <td width="18%" rowspan="2" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> 联系人<strong>:</strong></div></td>
                <td width="16%"><div align="right">
                    <input name="Sex" type="radio" value="1" <% if Sex="1" then response.Write("checked") end if%> />
                  先生</div></td>
                <td width="66%" colspan="2" rowspan="2"><input name="Receiver" type="text" class="form" id="name" size="30" value="<%=Receiver%>" />
                    <b></b></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td><div align="right">
                    <input type="radio" name="Sex" value="0" <% if Sex="0" then response.Write("checked") end if%> />
                  女士</div></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="30" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> 公司名称<strong>:</strong></div></td>
                <td height="30" colspan="3" bgcolor="#FFFFFF"><input type="text" name="CompanyName" size="40" value="<%=CompanyName%>" />
                  (4-100 字符) </td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="30" bgcolor="#FFFFFF"><div align="right">地址:</div></td>
                <td height="30" colspan="3"><input type="text" name="Add"  size="50" value="<%=Add%>" /></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="30" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> E-mail<strong>:</strong></div></td>
                <td height="30" colspan="3"><input type="text" name="Email"  size="30" value="<%=Email%>" />
                    <b></b></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="30" bgcolor="#FFFFFF"><div align="right"><font color="#FF0000">*</font> 公司电话<strong>:</strong></div></td>
                <td colspan="3"><input type="text" name="Phone" class="form" size="20" value="<%=Phone%>" />
                  (Tel:86-010-81991660)</td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="30" bgcolor="#FFFFFF"><div align="right">传真<strong>:</strong></div></td>
                <td colspan="3"><input type="text" name="Fax" size="20" value="<%=Fax%>" />
                  (Format:86-010-81991660)</td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="32" colspan="4" align="center"><input type="submit" name="nextstep" value="提交请求" />
                  &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
                  <input type="button" value="继续购物" language="javascript" onClick="javascript:window.close()" name="button" />
                  &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
                  <input name="reset" type="reset" id="reset" value="重新填写" />
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

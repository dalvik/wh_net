<!--#include file="Inc/SysProduct.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
UserName=session("UserName")			'��½�û�id
Receiver=request.form("Receiver")			'�����ֶ�
Sex=request.form("Sex")			'
Phone=request.form("Phone")		'�����ֶ�
Add=request.form("Add")	'�����ֶ�
Notes=request.form("Notes")			
Email=request.form("Email")			
Subject=request.form("Subject")
CompanyName=request.form("CompanyName")			
Fax=request.form("Fax")		
if UserName="" then 
    UserName="�ο�" 
end if
if Add="" then
	Add=null
end if
if Fax="" then
	Fax=null
end if


'�жϹ��ﳵ�Ƿ�Ϊ��
ProductList = Session("ProductList")
if productlist="" then
 response.redirect "error.asp?error=007"
 response.end
else
  sql_product="select * from Product where Product_Id in ("&productlist&") order by Product_Id"
  Set rs_order = conn.Execute(sql_product)
end if

BranchID="0022"
CoNo="000040"

'�������ڣ���ʽ��YYYYMMDD
yy=right(year(date),2)
mm=right("00"&month(date),2)
dd=right("00"&day(date),2)
riqi=yy & mm & dd

'���ɶ�������������Ԫ��,��ʽΪ��Сʱ�����ӣ���
xiaoshi=right("00"&hour(time),2)
fenzhong=right("00"&minute(time),2)
miao=right("00"&second(time),2)

'�����ⲿ���ڲ�������
BillNo=xiaoshi & fenzhong & miao
inBillNo=yy & mm & dd & "-" & xiaoshi & fenzhong & miao

Set rsadd=server.createobject("adodb.recordset")
rsadd.Open "select * from OrderList" ,conn,1,3
Set rsdetail=server.createobject("adodb.recordset")
rsdetail.Open "select * from OrderDetail" ,conn,1,3

'�����忪ʼ
'conn.Begintrans

'����֮һ��ʼд�붩���б���Ϣ
rsadd.AddNew 
rsadd("UserName")=UserName
rsadd("OrderNum")=inBillNo
rsadd("Receiver")=Receiver
rsadd("Sex")=Sex
rsadd("Phone")=Phone
rsadd("Add")=Add
rsadd("RecTime")=now()
if Subject<>"" then rsadd("Subject")=Subject
if Email<>"" then rsadd("Email")=Email
if CompanyName<>"" then rsadd("CompanyName")=CompanyName
if Fax<>"" then rsadd("Fax")=Fax
if Notes<>"" then rsadd("Notes")=Notes
if error>0 then
	response.write " ���������б����ɴ��󣡣�"
    return
end if
rsadd("Flag")="No"
rsadd.Update


While Not rs_order.EOF '�ѹ���Ĳ�Ʒ���϶�������д�붨����ϸ���ϱ���

 rsdetail.AddNew 
 rsdetail("UserName")=UserName		'�µ��û���
 rsdetail("OrderNum")=inBillNo		'��������
 rsdetail("Product_Id")=rs_order("Product_Id")		'��Ʒ����
 rsdetail("OrderTime")=date()

 IF ERROR>0 THEN
	response.write "����������ϸ��Ϣ�����ɴ��󣡣�"
	RETURN
 END if
 rsdetail.Update
 rs_order.MoveNext
Wend
'�����������
'conn.CommitTrans

rsdetail.close
set rsdetail=nothing
rsadd.close
set rsadd=nothing
rs_order.close
set rs_order=nothing
Session("ProductList") =""
%>
<style type="text/css">
<!--
.style1 {color: #6F6F6F}
-->
</style>
<style type="text/css">
<!--
.style2 {color: #CC0000}
-->
</style>
<!-- #include file="Head.asp" -->
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
  <table width="100%" height="212" border="0" align="center" cellpadding="0" cellspacing="1"  valign="absmiddle">
    <tr >
      <td height="64" align="center"><span class="style1">лл����ѯ���ύ�ɹ���������ס����ѯ�ۺ��룬�Ա��ѯ��</span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="100%" align="center" valign="middle"><br />
        ��л����ѯ���ǵĲ�Ʒ!<b><br />
          <br />
          </b><span class="style2">����ѯ�ۺ�����:</span><b><font color="B0266D"><b><%=inBillNo%></b><br />
            <br />
            </font> <br />
          </b></td>
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

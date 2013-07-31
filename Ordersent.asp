<!--#include file="Inc/SysProduct.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
UserName=session("UserName")			'登陆用户id
Receiver=request.form("Receiver")			'必填字段
Sex=request.form("Sex")			'
Phone=request.form("Phone")		'必填字段
Add=request.form("Add")	'必填字段
Notes=request.form("Notes")			
Email=request.form("Email")			
Subject=request.form("Subject")
CompanyName=request.form("CompanyName")			
Fax=request.form("Fax")		
if UserName="" then 
    UserName="游客" 
end if
if Add="" then
	Add=null
end if
if Fax="" then
	Fax=null
end if


'判断购物车是否为空
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

'交易日期，格式：YYYYMMDD
yy=right(year(date),2)
mm=right("00"&month(date),2)
dd=right("00"&day(date),2)
riqi=yy & mm & dd

'生成订单号所有所需元素,格式为：小时，分钟，秒
xiaoshi=right("00"&hour(time),2)
fenzhong=right("00"&minute(time),2)
miao=right("00"&second(time),2)

'产生外部和内部定单号
BillNo=xiaoshi & fenzhong & miao
inBillNo=yy & mm & dd & "-" & xiaoshi & fenzhong & miao

Set rsadd=server.createobject("adodb.recordset")
rsadd.Open "select * from OrderList" ,conn,1,3
Set rsdetail=server.createobject("adodb.recordset")
rsdetail.Open "select * from OrderDetail" ,conn,1,3

'事务定义开始
'conn.Begintrans

'操作之一开始写入订单列表信息
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
	response.write " 操作订单列表生成错误！！"
    return
end if
rsadd("Flag")="No"
rsadd.Update


While Not rs_order.EOF '把购买的产品资料读出来，写入定单详细资料表中

 rsdetail.AddNew 
 rsdetail("UserName")=UserName		'下单用户号
 rsdetail("OrderNum")=inBillNo		'订单号码
 rsdetail("Product_Id")=rs_order("Product_Id")		'产品编码
 rsdetail("OrderTime")=date()

 IF ERROR>0 THEN
	response.write "操作订单详细信息表生成错误！！"
	RETURN
 END if
 rsdetail.Update
 rs_order.MoveNext
Wend
'事务操作结束
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
  <table width="100%" height="212" border="0" align="center" cellpadding="0" cellspacing="1"  valign="absmiddle">
    <tr >
      <td height="64" align="center"><span class="style1">谢谢您，询价提交成功，请您记住您的询价号码，以便查询。</span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="100%" align="center" valign="middle"><br />
        感谢您咨询我们的产品!<b><br />
          <br />
          </b><span class="style2">您的询价号码是:</span><b><font color="B0266D"><b><%=inBillNo%></b><br />
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

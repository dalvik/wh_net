<!--#include file="Inc/SysProduct.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="Head.asp" -->
<%If Session("UserName")="" Then
response.redirect"UserServer.asp"
end if%>

<SCRIPT language=javascript>
<!--

function chkitem(str)
{	
var strSource ="0123456789-";
var ch;
var i;
var temp;

for (i=0;i<=(str.length-1);i++)
{

ch = str.charAt(i);
temp = strSource.indexOf(ch);
if (temp==-1)
{
return 0;
}
}
if (strSource.indexOf(ch)==-1)
{
return 0;
}
else
{
return 1;
}
}

function order_onsubmit()
{
if (chkitem(document.order.OrderNum.value)==0)
{
alert("请输入正确的询价号。");
document.order.OrderNum.focus();
return false;
}
}
//-->
</SCRIPT>


<body>

<DIV id=index_container>
  <div id="index_left">
    <div id="left">
      <div id="left_title">客户服务</div>
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


  <LI>
  <a href="UserEdit.asp">修改会员资料</a>
    </li>
  
    <LI>
  <a href="UserEditPwd.asp">修改会员密码</a>
  </li>
     <LI>
  <a href="UserOrder.asp">购物订单查询</a>
  </li>
    <LI>
  <a href="FeedbackMember.asp">站内留言中心</a>
  </li>
    <LI>
  <a href="UserLogout.asp?Action=Ch">退出会员中心</a>
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
  <DIV id=right_title>用户中心</DIV>
<DIV class=text>
  <table width=100% height="136" align="center" cellpadding=4 cellspacing=1>
                <tbody>
                  <tr valign=top >
                    <td  width="100%" height="18"><form action="orderofind.asp" method="post" name=order onSubmit="return order_onsubmit()" language=javascript>
                        <p align="left" style="margin-bottom: 0"> <font color="#FF0000">亲爱的：&nbsp;</font><font size="3" face="Arial"><%=Session("UserName")%></font>
                            <%rs.close%>
                            <%


UserName=Session("UserName")
set Rs = Server.CreateObject("ADODB.recordset")
sql="select * from OrderList where UserName='"&UserName&"'"
rs.open sql,conn,1,1


%>
                        </p>
                        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">　　订单详细资料查询，请输入您要查询的询价号码
                            <input type="text" name="OrderNum" size="16">
                            <input type="submit" value="查询" name="B1">
                        </form>
                        <div align="center">
                          <table cellspacing=0 cellpadding=0 width=100%>
                            <tbody>
                              <tr>
                                <td align=middle width="100%"><p align="left" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><br />
                                  询价列表:<br />
                                  <br />
                                </p>
                                    <table border="0" cellpadding="2" cellspacing="1" width="100%" bgcolor="#CCCCCC">
                                      <tr bgcolor="#E0E0E0" >
                                        <td width="26%" height="45" align="center" bgcolor="#EFEFEF"><span class="style1">询价编号</span></td>
                                        <td width="21%" height="26" align="center" bgcolor="#EFEFEF"><span class="style1">客户姓名</span></td>
                                        <td width="53%" height="26" align="center" bgcolor="#EFEFEF"><span class="style1">对您的询价的处理情况</span></td>
                                      </tr>
                                      <% While Not rs.EOF %>
                                      <tr >
                                        <td width="26%" height="45" align="center" bgcolor="#FFFFFF"><%response.Write "&nbsp;<a href='#' onclick=""javascript:window.open('Orderofind.asp?OrderNum=" & rs("OrderNum") &"', 'newwindow', 'height=550, width=550, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rs("OrderNum") & "'><font color='#FF0000'>" &rs("OrderNum") & "</font></a>"%>
                                        </td>
                                        <td width="21%" align="center" height="25" bgcolor="#FFFFFF"><%=rs("Receiver")%></td>
                                        <td width="53%" align="center" height="25" bgcolor="#FFFFFF"><%If rs("Flag")="Yes" Then%>
                        已处理
                          <%else%>
                          <font color="#FF0000"> 未处理</font>
                          <%End If%>
                                        </td>
                                      </tr>
                                      <%
rs.MoveNext
Wend
rs.close
%>
                                  </table></td>
                              </tr>
                            </tbody>
                          </table>
                      </div></td>
                  </tr>
                  <tr valign=top >
                    <td  width="100%" height="15">&nbsp;&nbsp; </td>
                  </tr>
                </tbody>
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
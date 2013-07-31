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
      <div id="left_title">留言中心</div>
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
<a href="FeedbackView.asp">查看留言</a>    </li>
  
    <LI>
<a href="Feedback.asp">我要留言</a>  </li>
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
  <DIV id=right_title>留言反馈</DIV>
<DIV class=text>
  <form method="post" action="FeedbackSave.asp?Language=ch">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>尊敬的客户：
          <p>　　如果您对我们的产品或服务有任何意见和建议请及时告诉我们，我们将尽快给您满意的答复。<br />
            如果您注册一个会员号，以后每次留言时只要登录再不用重复填写你的联系信息了，并且通过会员管理可集中查看只属于您的留言。（非注册会员也可直接留言） <br />
            <br />
          </p></td>
      </tr>
    </table>
    <table width="100%" height="409"
border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E3E3E3">
      <tr>
        <input type="hidden" name="Username" value="<%=Username%>" />
        <td height="25" align="right" bgcolor="#FFFFFF">主题： </td>
        <td height="25" bgcolor="#FFFFFF"><input type="text" name="Title" size="42" maxlength="36" style="font-size: 14px">
          *</td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">内容 *：</td>
        <td height="25" bgcolor="#FFFFFF"><textarea rows="10" name="Content" cols="45" style="font-size: 14px" ></textarea></td>
      </tr>
      <tr>
        <td width="23%" height="25" align="right" bgcolor="#FFFFFF">公司名称：</td>
        <td width="77%" height="-6" bgcolor="#FFFFFF"><font>
          <input type="text" name="CompanyName" size="30" maxlength="36" value="<%=rs("CompanyName")%>" style="font-size: 14px">
        </font>* </td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">公司地址：</td>
        <td height="-2" bgcolor="#FFFFFF"><font>
          <input name="Add" type="text"  id="Add" style="font-size: 14px" value="<%=rs("Add")%>" size="40" maxlength="60" />
        </font></td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">邮编：</td>
        <td height="-2" bgcolor="#FFFFFF"><font>
          <input name="Postcode" type="text"  id="Postcode" style="font-size: 14px" value="<%=rs("Postcode")%>" size="12" maxlength="6" />
        </font></td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">联系人：</td>
        <td width="77%" height="-2" bgcolor="#FFFFFF"><font>
          <input type="text" name="Receiver" size="12" maxlength="30" value="<%=rs("Receiver")%>" style="font-size: 14px">
        </font> * </td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">联系电话：</td>
        <td width="77%" height="-1" bgcolor="#FFFFFF"><font>
          <input type="text" name="Phone" size="24" maxlength="36" value="<%=rs("Phone")%>" style="font-size: 14px">
        </font>* </td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">手机：</td>
        <td height="11" bgcolor="#FFFFFF"><font>
          <input name="Mobile" type="text" id="Mobile" style="font-size: 14px" value="<%=rs("Mobile")%>" size="24" maxlength="36" />
        </font></td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">联系传真：</td>
        <td width="77%" height="11" bgcolor="#FFFFFF"><font>
          <input type="text" name="Fax" size="18" maxlength="36" value="<%=rs("Fax")%>" style="font-size: 14px">
        </font></td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">E-mail：</td>
        <td height="11" bgcolor="#FFFFFF"><font>
          <input type="text" name="Email" size="18" maxlength="36" value="<%=rs("Email")%>" style="font-size: 14px">
        </font></td>
      </tr>
      <tr>
        <td height="25" align="right" bgcolor="#FFFFFF">悄悄话：</td>
        <td width="77%" height="11" bgcolor="#FFFFFF"><% if session("username")<>"" then  %>
            <input type="radio" name="Publish" value="1" />
          是
          <input name="Publish" type="radio" value="0" checked="checked" />
          否
          <% else %>
          <input name="Publish" type="hidden" value="0" checked="checked" />
          <%	  response.write "<font color='#ff6600'>非注册会员不可设</font>"
					end if%>
          (只有管理员和自己能看)</td>
      </tr>
      <tr>
        <td width="23%" height="0" valign="top" bgcolor="#FFFFFF"></td>
        <td width="77%" height="0" valign="top" bgcolor="#FFFFFF"><input type="submit" value="提交留言"
name="cmdOk" />
            <input type="reset" value="重写" name="cmdReset" />
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

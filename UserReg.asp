<!--#include file="Inc/SysProduct.asp" -->
<!-- #include file="Head.asp" -->
<script language=javascript>
function checkreg()
{
  if (document.UserReg.UserName.value=="")
	{
	alert("�������û�����");
	document.UserReg.UserName.focus();
	return false;
	}
  else
    {
	document.reg.username.value=document.UserReg.UserName.value;
	var popupWin = window.open('UserCheckReg.asp', 'CheckReg', 'scrollbars=no,width=340,height=200');
	document.reg.submit();
	}
}
</script>
<body>
<DIV id=index_container>
  <div id="index_left">
    <div id="left">
      <div id="left_title">�ͻ�����</div>
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
  <a href="UserEdit.asp">�޸Ļ�Ա����</a>
    </li>
  
    <LI>
  <a href="UserEditPwd.asp">�޸Ļ�Ա����</a>
  </li>
     <LI>
  <a href="UserOrder.asp">���ﶩ����ѯ</a>
  </li>
    <LI>
  <a href="FeedbackMember.asp">վ����������</a>
  </li>
    <LI>
  <a href="UserLogout.asp?Action=Ch">�˳���Ա����</a>
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
  <DIV id=right_title>�û�ע��</DIV>
<DIV class=text>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="1"><form action='UserRegPost.asp' method='post' name='UserReg' id="UserReg">
          <table width="95%" border="0" align="center" cellpadding="5" cellspacing="1" bordercolor="#FFFFFF" style="border-collapse: collapse">
            <tr align="center">
              <td height="20" colspan="2"><b>���û�ע��</b></td>
            </tr>
            <tr>
              <td width="40%"><b>�û�����</b><br />
                ����С��4���ַ���2�����֣�</td>
              <td width="60%"><input   maxlength="14" size="30" name="UserName" />
                  <font color="#FF0000">*</font>
                  <input name="Check" type="button" id="Check" value="����û���" onClick="checkreg();" /></td>
            </tr>
            <tr>
              <td width="40%"><b>����(����6λ)��</b><br />
                ���������룬���ִ�Сд�� ��Ҫʹ������ '*'��' '�������ַ�</td>
              <td width="60%"><input   type="password" maxlength="12" size="30" name="Password" />
                  <font color="#FF0000">*</font> </td>
            </tr>
            <tr>
              <td width="40%"><strong>ȷ������(����6λ)��</strong><br />
              </td>
              <td width="60%"><input   type="password" maxlength="12" size="30" name="PwdConfirm" />
                  <font color="#FF0000">*</font> </td>
            </tr>
            <tr>
              <td width="40%"><strong>�������⣺</strong><br />
                �����������ʾ����</td>
              <td width="60%"><input   type="text" maxlength="50" size="30" name="Question" />
                  <font color="#FF0000">*</font> </td>
            </tr>
            <tr>
              <td width="40%"><strong>����𰸣�</strong><br />
                �����������ʾ����𰸣�����ȡ������</td>
              <td width="60%"><input   type="text" maxlength="20" size="30" name="Answer" />
                  <font color="#FF0000">*</font> </td>
            </tr>
            <tr>
              <td width="40%"><strong>�Ա�</strong><br />
                ��ѡ�������Ա�</td>
              <td width="60%"><input type="radio" checked="checked" value="1" name="sex" />
                �� &nbsp;&nbsp;
                <input type="radio" value="0" name="sex" />
                Ů</td>
            </tr>
            <tr>
              <td width="40%"><strong>Email��ַ��</strong><br />
                ��������Ч���ʼ���ַ</td>
              <td width="60%"><input   maxlength="50" size="30" name="Email" />
                  <font color="#FF0000">*</font></td>
            </tr>
            <tr>
              <td><strong>��˾��ַ��</strong></td>
              <td width="60%"><input name="homepage" id="homepage" value="http://" size="30"   maxlength="50" /></td>
            </tr>
            <tr>
              <td width="40%"><strong>��˾���ƣ�</strong><br />
                ���Ĺ�˾����</td>
              <td width="60%"><input name="CompanyName" id="CompanyName" size="30"   maxlength="100" /></td>
            </tr>
            <tr>
              <td><strong>�ջ���ַ��</strong></td>
              <td><input name="Add" id="Add" size="30"   maxlength="100" />
                  <font color="#FF0000">*</font></td>
            </tr>
            <tr>
              <td><strong>�ջ��ˣ�</strong></td>
              <td><input name="Receiver" id="Receiver" size="30"   maxlength="100" /></td>
            </tr>
            <tr>
              <td><strong>�������룺</strong></td>
              <td width="60%"><input name="postcode" id="postcode" size="30" maxlength="20" />
                  <font color="#FF0000">*</font></td>
            </tr>
            <tr>
              <td><strong>��ϵ�绰��<br />
              </strong>��ʽ001-81991660<strong> </strong></td>
              <td width="60%"><input name="Phone" id="Phone" size="30" maxlength="20" />
                  <font color="#FF0000">*</font></td>
            </tr>
            <tr>
              <td><strong>�ֻ���</strong></td>
              <td><input name="Mobile" id="Mobile" size="30" maxlength="20" /></td>
            </tr>
            <tr>
              <td width="40%"><strong>�� �棺</strong></td>
              <td width="60%"><input name="Fax" id="Fax" size="30" maxlength="50" /></td>
            </tr>
          </table>
        <div align="center">
            <input   type="submit" value=" ע �� " name="Submit2" />
          &nbsp;
          <input name="Reset"   type="reset" id="Reset" value=" �� �� " />
          </div>
      </form>
          <form action='UserCheckreg.asp' method='post' name='reg' target='CheckReg' id="reg">
            <input type='hidden' name='username' value='' />
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

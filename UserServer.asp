<!--#include file="Inc/conn.asp" -->
<!--#include file="Inc/Function.asp" -->
<!--#include file="Inc/Config.asp" -->
<!-- #include file="Head.asp" -->
<!--#include file="inc/md5.asp"-->
<script language=javascript>
	function CheckForm()
	{
		if(document.UserLogin.UserName.value=="")
		{
			alert("�������û�����");
			document.UserLogin.UserName.focus();
			return false;
		}
		if(document.UserLogin.Password.value == "")
		{
			alert("���������룡");
			document.UserLogin.Password.focus();
			return false;
		}
	}
</script>
<%
Action=Trim(request("Action"))
if Action="Login" then
 dim sql
 dim rs
 dim username
 dim password
 username=replace(trim(request("username")),"'","")
 password=replace(trim(Request("password")),"'","")
 password=md5(password)
 set rs=server.createobject("adodb.recordset")
 sql="select * from [User] where LockUser=False and username='" & username & "' and password='" & password &"'"
 rs.open sql,conn,1,3
 if not(rs.bof and rs.eof) then
	if password=rs("password") then
	    rs("LoginIP")=Request.ServerVariables("REMOTE_ADDR")
		rs("LastLoginTime")=now()
		rs("logins")=rs("logins")+1
		rs.update
		session("UserName")=rs("username")		
		Response.Redirect "UserServer.asp"
	end if
 end if
 rs.close
 conn.close
 set rs=nothing
 set conn=nothing
end if
%>

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
  <DIV id=right_title>�û�����</DIV>
<DIV class=text><table width="100%" height="50%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="206" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <TR>
                        <TD><%If Session("UserName")="" Then%>
                            <table width="100%" border="0" cellpadding="0" cellspacing="3">
                              <tr>
                                <td colspan="2" align="right"><p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p></td>
                              </tr>
                              <tr>
                                <td colspan="2" align="right"><div align="center">
                                    <table class="main" cellSpacing="0" cellPadding="2" width="100%" border="0" height="1">
                                      <tbody>
                                        <tr>
                                          <td width="100%" height="1"><table class='border'  width='93%' border='0' align="center" cellpadding='5' cellspacing='0'>
                                              <form action='UserServer.asp?Action=Login' method='post' name='UserLogin' onSubmit='return CheckForm();'>
                                                <tr>
                                                  <td height='35' colspan="2" align='left'><div align="center" class="STYLE1">���޷������Ա���ģ����ȵ�¼��������������ǵĻ�Ա��������<font color="#006699"><a href="UserReg.asp"><font color="#FF0000">ע��</font></a>��</font></div></td>
                                                </tr>
                                                <tr>
                                                  <td width="45%" height='35' align='right'>�û�����</td>
                                                  <td width="55%" height='35' align="left"><input name='UserName' type='text' id='UserName' size='15' maxlength='20'></td>
                                                </tr>
                                                <tr>
                                                  <td height='35' align='right'>��&nbsp;&nbsp;&nbsp;&nbsp;�룺</td>
                                                  <td height='35' align="left"><input name='Password' type='password' id='Password' size='15' maxlength='20'></td>
                                                </tr>
                                                <tr align='center'>
                                                  <td height='35' colspan='2'><input name='Login' type='submit' id='Login' value=' ��¼ '>
                                                      <input name='Reset' type='reset' id='Reset' value=' ��� '></td>
                                                </tr>
                                                <tr>
                                                  <td height='35' colspan='2' align='center'><a href='UserReg.asp' target='_blank'>�û�ע��</a>&nbsp;&nbsp;<a href='GetPassword.asp' target='_blank'>��������</a></td>
                                                </tr>
                                              </form>
                                          </table></td>
                                        </tr>
                                      </tbody>
                                    </table>
                                </div></td>
                              </tr>
                              <tr>
                                <td height="19" colspan="2" align="right">&nbsp;</td>
                              </tr>
                              <tr>
                                <td width="20%" align="right"></td>
                                <td width="80%"><span style="font-size: 10pt; font-weight: bold; color: #FF9900">���ǵĻ�Ա�����¹��ܣ�</span></td>
                              </tr>
                              <tr>
                                <td width="20%" height="30" align="right">&nbsp;</td>
                                <td width="80%" height="20">1���޸����Ļ�Աע�����ϣ�</td>
                              </tr>
                              <tr>
                                <td width="20%" height="30" align="right">&nbsp;</td>
                                <td width="80%" height="20"><p>2����ѯ����ѯ�ۣ��Լ�ѯ�۵Ĵ��������</p></td>
                              </tr>
                              <tr>
                                <td width="20%" height="30" align="right">&nbsp;</td>
                                <td width="80%" height="20">3��ר��˽�����Բ��������ڴ˸��������ԺͲ鿴���ǵĻظ���</td>
                              </tr>
                          </table>
                            <%else%>
                            <TABLE width="85%" align="center" cellPadding=0 cellSpacing=0>
                              <TBODY>
                                <TR vAlign=top >
                                  <TD  width="100%"><p style="margin-bottom: 0"> <br>
                      �װ���<%=Session("UserName")%>��</p>
                                      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp; </p>
                                      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"> �����������Ѿ������Ա�������ģ�����ֻ��ע���Ա���ܷ��ʡ������������޸�����ע����Ϣ�����������ԡ��鿴���Ƕ������ԵĴ𸴣�Ҳ���Բ�ѯ����ѯ�ۼ�ѯ�۴��������</p></TD>
                                </TR>
                              </TBODY>
                            </TABLE>
                            <%end if%>
                        </TD>
                      </TR>
                  </table></td>
                </tr>
              </table></DIV>
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

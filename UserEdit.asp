<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim Action,UserName,FoundErr,ErrMsg
dim rsUser,sqlUser
Action=trim(request("Action"))
UserName=trim(request("UserName"))
if UserName="" then
	UserName=session("UserName")
end if
if  UserName="" then
	if Action="" then
		response.redirect "UserServer.asp"
	else
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
	end if
end if
if FoundErr=true then
	call WriteErrMsg()
else
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from [User] where UserName='" & UserName & "'"
	rsUser.Open sqlUser,conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
		call writeErrMsg()
	else
		if Action="Modify" then
			dim Sex,Email,Homepage
			Sex=trim(Request("Sex"))
			Email=trim(request("Email"))
			Homepage=trim(request("Homepage"))
			CompanyName=trim(request("CompanyName"))
			Add=trim(request("Add"))
			Receiver=trim(request("Receiver"))
			Postcode=trim(request("Postcode"))
			Phone=trim(request("Phone"))
			Mobile=trim(request("Mobile"))
			Fax=trim(request("Fax"))
			if Sex="" then
				founderr=true
				errmsg=errmsg & "<br><li>�Ա���Ϊ��</li>"
			else
				sex=cint(sex)
				if Sex<>0 and Sex<>1 then
					Sex=1
				end if
			end if
			if Email="" then
				founderr=true
				errmsg=errmsg & "<br><li>Email����Ϊ��</li>"
			else
				if IsValidEmail(Email)=false then
					errmsg=errmsg & "<br><li>����Email�д���</li>"
			   		founderr=true
				end if
			end if
			if CompanyName="" then
				founderr=true
				errmsg=errmsg & "<br><li>��˾���Ʋ���Ϊ��</li>"
			end if
			if Add="" then
				founderr=true
				errmsg=errmsg & "<br><li>�ջ���ַ����Ϊ��</li>"
			end if
			if Receiver="" then
				founderr=true
				errmsg=errmsg & "<br><li>�ջ��˲���Ϊ��</li>"
			end if
			if Phone="" then
				founderr=true
				errmsg=errmsg & "<br><li>��ϵ�绰����Ϊ��</li>"
			end if
			if FoundErr<>true then
				rsUser("Sex")=Sex
				rsUser("Email")=Email
				rsUser("HomePage")=HomePage
				rsUser("CompanyName")=CompanyName
				rsUser("Add")=Add
				rsUser("Receiver")=Receiver
				rsUser("Postcode")=Postcode
				rsUser("Phone")=Phone
				rsUser("Mobile")=Mobile
				rsUser("Fax")=Fax				
				rsUser("LastLoginTime")=Now()
				rsUser.update
				response.write"<SCRIPT language=JavaScript>alert('��Ա�����޸ĳɹ���');"
                response.write"javascript:history.go(-1)</SCRIPT>"				
			else
				call WriteErrMsg()
			end if
		else

%>
<!-- #include file="Head.asp"-->
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
<DIV class=text>
  <form action="UserEdit.asp" method="post" name="Form1" id="Form1">
    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E3E3E3" class='border'>
      <tr align="center" bgcolor="#FFFFFF" >
        <td height="40" colspan="2"><font color="#FF6600"><b>�޸�ע���û���Ϣ</b></font><font color="#ffffff" class="font14">&nbsp;</font></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="267" height="30" align="right">�� �� ����</td>
        <td width="397" height="30"><input name="UserName" value="<%=rsUser("UserName")%>" size="30"   maxlength="14" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="267" height="30" align="right">�Ա�</td>
        <td height="30"><input type="radio" value="1" name="sex" <%if rsUser("Sex")=1 then response.write "CHECKED"%> />
          �� &nbsp;&nbsp;
          <input type="radio" value="0" name="sex" <%if rsUser("Sex")=0 then response.write "CHECKED"%> />
          Ů</td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="267" height="30" align="right">Email��ַ��</td>
        <td height="30"><input name="Email" value="<%=rsUser("Email")%>" size="30"   maxlength="50" />
        </td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="267" height="30" align="right">��ҳ��</td>
        <td height="30"><input   maxlength="100" size="30" name="homepage" value="<%=rsUser("HomePage")%>" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="267" height="30" align="right">��˾���ƣ�</td>
        <td height="30"><input name="CompanyName" value="<%=rsUser("CompanyName")%>" size="30" maxlength="20" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="267" height="30" align="right">�ջ���ַ��</td>
        <td height="30"><input name="Add" value="<%=rsUser("Add")%>" size="30" maxlength="50" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="30" align="right">�ջ��ˣ�</td>
        <td height="30"><input name="Receiver" value="<%=rsUser("Receiver")%>" size="30" maxlength="50" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="30" align="right">�������룺</td>
        <td height="30"><input name="Postcode" value="<%=rsUser("Postcode")%>" size="30" maxlength="50" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="30" align="right">��ϵ�绰��<br /></td>
        <td height="30"><input name="Phone" value="<%=rsUser("Phone")%>" size="30" maxlength="50" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="30" align="right">�ֻ���</td>
        <td height="30"><input name="Mobile" value="<%=rsUser("Mobile")%>" size="30" maxlength="50" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="30" align="right">�� �棺</td>
        <td height="30"><input name="Fax" value="<%=rsUser("Fax")%>" size="30" maxlength="50" /></td>
      </tr>
      <tr align="center" bgcolor="#FFFFFF" >
        <td height="40" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify" />
            <input name="Submit"   type="submit" id="Submit" value="�����޸Ľ��" />
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

</BODY></HTML>
<%
		end if
	end if
	rsUser.close
	set rsUser=nothing
end if
call CloseConn()
%>
<!--#include file="Inc/SysProduct.asp" -->
<!--#include file="Inc/md5.asp"-->
<!-- #include file="Head.asp" -->
<%
dim UserName,Password,PwdConfirm,Question,Answer,Sex,Email,HomePage,CompanyName,Add,Receiver,Postcode,Phone,Fax
UserName=trim(request("UserName"))
Password=trim(request("Password"))
PwdConfirm=trim(request("PwdConfirm"))
Question=trim(request("Question"))
Answer=trim(request("Answer"))
Sex=trim(Request("Sex"))
Email=trim(request("Email"))
HomePage=trim(request("HomePage"))
CompanyName=trim(request("CompanyName"))
Add=trim(request("Add"))
Receiver=trim(request("Receiver"))
Postcode=trim(request("Postcode"))
Phone=trim(request("Phone"))
Mobile=trim(request("Mobile"))
Fax=trim(request("Fax"))
if UserName="" or strLength(UserName)>14 or strLength(UserName)<4 then
	founderr=true
	errmsg=errmsg & "<br><li>�������û���(���ܴ���14С��4)</li>"
else
  	if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"��")>0 or Instr(UserName,"$")>0 then
		errmsg=errmsg+"<br><li>�û����к��зǷ��ַ�</li>"
		founderr=true
	end if
end if
if Password="" or strLength(Password)>12 or strLength(Password)<6 then
	founderr=true
	errmsg=errmsg & "<br><li>����������(���ܴ���12С��6)</li>"
else
	if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"��")>0 or Instr(Password,"$")>0 then
		errmsg=errmsg+"<br><li>�����к��зǷ��ַ�</li>"
		founderr=true
	end if
end if
if PwdConfirm="" then
	founderr=true
	errmsg=errmsg & "<br><li>������ȷ������(���ܴ���12С��6)</li>"
else
	if Password<>PwdConfirm then
		founderr=true
		errmsg=errmsg & "<br><li>�����ȷ�����벻һ��</li>"
	end if
end if
if Question="" then
	founderr=true
	errmsg=errmsg & "<br><li>������ʾ���ⲻ��Ϊ��</li>"
end if
if Answer="" then
	founderr=true
	errmsg=errmsg & "<br><li>����𰸲���Ϊ��</li>"
end if
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
if Add="" then
	founderr=true
	errmsg=errmsg & "<br><li>�ջ���ַ����Ϊ��</li>"
end if
if Postcode="" then
	founderr=true
	errmsg=errmsg & "<br><li>�������벻��Ϊ��</li>"
end if
if Phone="" then
	founderr=true
	errmsg=errmsg & "<br><li>��ϵ�绰����Ϊ��</li>"
end if

if founderr=false then
	dim sqlReg,rsReg
	sqlReg="select * from [User] where UserName='" & Username & "'"
	set rsReg=server.createobject("adodb.recordset")
	rsReg.open sqlReg,conn,1,3
	if not(rsReg.bof and rsReg.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>��ע����û��Ѿ����ڣ��뻻һ���û��������ԣ�</li>"
	else
		rsReg.addnew
		rsReg("UserName")=UserName
		rsReg("Password")=md5(Password)
		rsReg("Question")=Question
		rsReg("Answer")=md5(Answer)
		rsReg("Sex")=Sex
		rsReg("Email")=Email
		rsReg("HomePage")=HomePage
		rsReg("CompanyName")=CompanyName
		rsReg("Add")=Add
		rsReg("Receiver")=Receiver
		rsReg("Postcode")=Postcode
		rsReg("Phone")=Phone
		rsReg("Mobile")=Mobile
		rsReg("Fax")=Fax
		rsReg("RegDate")=Now()
		rsReg("LockUser")=false
		rsReg.update
		founderr=false
	end if
	rsReg.close
	set rsReg=nothing
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
<DIV class=text>

  <table width="650" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="56" align="center">
        <%
if founderr=false then
	call RegSuccess()
else
	call WriteErrmsg()
end if
%></td>
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



<!-- #include file="Inc/Foot.asp" -->
</BODY></HTML>

<%call CloseConn
sub WriteErrMsg()
    response.write "<table align='center' width='650' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='15'>�������µ�ԭ����ע���û���</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>" & errmsg & "<p align='center'>��<a href='javascript:onclick=history.go(-1)'>�� ��</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub

sub RegSuccess()
    response.write "<table align='center' width='650' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='15'>�ɹ�ע���û���</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>��ע����û�����" & UserName & "<p align='center'>��<a href='javascript:onclick=window.close()'>�� ��</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub
%>

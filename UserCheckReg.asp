<!--#include file="Inc/conn.asp"-->
<!--#include file="Inc/function.asp"-->
<%
dim UserName,FoundErr,ErrMsg

UserName=trim(request("UserName"))
Response.Write(UserName)
if UserName="" or strLength(UserName)>14 or strLength(UserName)<4 then
	founderr=true
	errmsg=errmsg & "<br><li>�������û���(���ܴ���14С��4)</li>"
else
  	if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"��")>0 or Instr(UserName,"$")>0 then
		errmsg=errmsg+"<br><li>�û����к��зǷ��ַ�</li>"
		founderr=true
	end if
end if
if founderr=false then
	dim sqlCheckReg,rsCheckReg
	sqlCheckReg="select * from [User] where UserName='" & Username & "'"
	set rsCheckReg=server.createobject("adodb.recordset")
	rsCheckReg.open sqlCheckReg,conn,1,1
	if not(rsCheckReg.bof and rsCheckReg.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>��" & UserName & "���Ѿ����ڣ��뻻һ���û��������ԣ�</li>"
	else
	end if
	rsCheckReg.close
	set rsCheckReg=nothing
end if		
%>
<html>
<head>
<title>����û���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="mt_style.css" type=text/css rel=stylesheet>
</head>

<body>
<br>
<%
if founderr=false then
	call Success()
else
	call WriteErrmsg()
end if
%>
</body>
</html>
<%
call CloseConn

sub WriteErrMsg()
    response.write "<table align='center' width='300' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='15'>������ʾ</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'>" & errmsg & "<p align='center'>��<a href='javascript:onclick=window.close()'>�� ��</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub

sub Success()
    response.write "<table align='center' width='300' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='15'>��ϲ�㣡</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>��" & UserName & "����δ����ʹ�ã��Ͻ�ע��ɣ�<p align='center'>��<a href='javascript:onclick=window.close()'>�� ��</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim sql
dim rs
dim username
dim password
username=replace(trim(request("username")),"'","")
password=replace(trim(Request("password")),"'","")
password=md5(password)
set rs=server.createobject("adodb.recordset")
sql="select * from [User] where LockUser=False and username='" & username & "' and password='" & password &"'"
rs.open sql,conn,1,1
if not(rs.bof and rs.eof) then
	if password=rs("password") then
		session("UserName")=rs("username")
		'session("purview")=999
		Response.Redirect "index.asp"
	end if
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
<html>
<head>
<style type="text/css">
A:link    {	 color: #333333;TEXT-DECORATION: none	}
A:visited {	 color: #333333;TEXT-DECORATION: none	}
A:active  {	 color: #003300;TEXT-DECORATION: none	}
A:hover   {	 color: #003300;TEXT-DECORATION: underline overline	}
.navtrail {  COLOR: #eeeeee; FONT-SIZE: 12px; LINE-HEIGHT: 12px }
A.navtrail:link {  COLOR: #eeeeee; CURSOR: w-resize }
A.navtrail:visited {  COLOR: #eeeeee; CURSOR: w-resize }
A.navtrail:active {  COLOR: #eeeeee; CURSOR: w-resize }
A.navtrail:hover {  COLOR: #eeeeee; CURSOR: e-resize }
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #cccccc; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #cccccc; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #cccccc; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #cccccc}
<!--
td {  font-family: "宋体"; font-size: 9pt; color: #333333; text-decoration: none}
-->
</style>
</head>
<body>
<br><br><br>
<table align='center' width='300' border='0' cellpadding='4' cellspacing='0' class='border'>
	<tr >
		
    <td height='15' colspan='2' align="center" class='title'>操作: 确认身份失败!</td>
</tr>
<tr>
    <td height='23' colspan='2' align="center" class='tdbg'> <br>
      <br>
      用户名或密码错误！或此帐号被锁定！！<br>
      <br> <a href='javascript:onclick=history.go(-1)'>【返回】</a> <br>
      <br></td>
</tr></table>
</body></html>
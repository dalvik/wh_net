<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
option explicit
response.buffer=true	
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="admin/admin.asp"-->
<!--#include file="inc/function.asp"-->
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<style type="text/css">body {font-size:	9pt}</style>
</head>
<BODY bgcolor="#FFFFFF" MONOSPACE>
<%
dim Action,FoundErr,ErrMsg
dim ID,sql,rs
Action=trim(request("Action"))
ID=trim(request("ID"))

select case Action
 case "Modify"
	if ID="" then
		response.write "请指定要修改的产品ID"
	else
		ID=Clng(ID)
		sql="select * from Product where ID=" & ID & ""
		Set rs= Server.CreateObject("ADODB.Recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			response.write "找不到产品"
		else
			response.write "<p>" & rs("Content") & "</p>"
		end if
		rs.close
		set rs=nothing 	
   end if
 case "News"
   if ID="" then
		response.write "请指定要修改的新闻ID"
	else
		ID=Clng(ID)
		sql="select * from News where ID=" & ID & ""
		Set rs= Server.CreateObject("ADODB.Recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			response.write "找不到新闻"
		else
			response.write "<p>" & rs("Content") & "</p>"
		end if
		rs.close
		set rs=nothing 	
   end if
   case "Down"
   if ID="" then
		response.write "请指定要修改的下载ID"
	else
		ID=Clng(ID)
		sql="select * from Download where ID=" & ID & ""
		Set rs= Server.CreateObject("ADODB.Recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			response.write "找不到下载"
		else
			response.write "<p>" & rs("Content") & "</p>"
		end if
		rs.close
		set rs=nothing 	
   end if  
   
end select
%>
</body>
</html>
<%
call CloseConn()
%>
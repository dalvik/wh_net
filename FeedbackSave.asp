<!--#include file="Inc/Conn.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<%
If request.form("CompanyName")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û���빫˾���ƣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("Receiver")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û������ϵ�ˣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("Phone")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û������ϵ�绰���뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("title")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û������⣬�뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("content")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û�����������ݣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
Add=trim(request.form("Add"))
Postcode=trim(request.form("Postcode"))
Mobile=trim(request.form("Mobile"))
Fax=trim(request.form("Fax"))
email=trim(request.form("email"))
Language=trim(request.QueryString("Language"))

if Add="" then
	Add=null
end if
if Postcode="" then
	Postcode=null
end if
if Mobile="" then
	Mobile=null
end if
if Fax="" then
	Fax=null
end if
if Email="" then
	Email=null
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Feedback"
rs.open sql,conn,1,3

rs.addnew
if session("username")="" then
  rs("Username")="δע���û�"
else
  rs("Username")=trim(request.form("Username"))
end if
rs("CompanyName")=trim(request.form("CompanyName"))
rs("Add")=Add
rs("Postcode")=Postcode
rs("Receiver")=trim(request.form("Receiver"))
rs("Phone")=trim(request.form("Phone"))
rs("Mobile")=Mobile
rs("Fax")=Fax
rs("email")=email
rs("title")=trim(request.form("title"))
rs("content")=htmlencode2(request.form("content"))
rs("Publish")=trim(request.form("Publish"))
if Language="ch" then 
 rs("Language")="0"
else
 rs("Language")="1"
end if 
rs("time")=date()
rs.update
rs.close

if Language="ch" then 
 response.redirect "FeedbackView.asp"
else
 response.redirect "EnFeedbackView.asp"
end if 

%>
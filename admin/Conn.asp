<%
dim conn
dim connstr
dim db
db="../Databases/%#@$@#FDS@#$%%#.asp" '���ݿ��ļ�λ��
on error resume next
connstr="DBQ="+server.mappath(""&db&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
if err then
err.clear
else
conn.open connstr
end if
function jincheng (p)
jincheng=p-100000000000000
end function
sub CloseConn()
	conn.close
	set conn=nothing
end sub

%>
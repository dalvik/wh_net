<%@language=vbscript codepage=936 %>
<!--#include file="Conn.asp"-->
<!--#include file="Admin.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim strFileName
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages
dim i,j
dim ID
dim Title
dim sql,rs
dim BigClassName,SmallClassName

BigClassName=trim(request("BigClassName"))
SmallClassName=trim(request("SmallClassName"))
strFileName="Down_Manage.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName

Title=Trim(request("Title"))
ID=Request("ID")

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

sql="select * from download where ID>0"
if BigClassName<>"" then
	sql=sql & "and BigClassName='" & BigClassName & "' "
	if SmallClassName<>"" then
		sql=sql & " and SmallClassName='" & SmallClassName & "' "
	end if
end if
if Title<>"" then
	sql=sql & " and Title like '%" & Title & "%' "
end if
sql=sql & " order by id desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.del.chkAll.checked){
	document.del.chkAll.checked = document.del.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
  }
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ��ѡ�еĲ�Ʒ��һ��ɾ�������ָܻ���"))
     return true;
   else
     return false;
	 
}

</SCRIPT>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="862" align="center" valign="top"> <br>
      <p align="center"><strong>�� 
        �� �� ��</strong></p>
      <table width="620" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#000000" class="border">
        <tr class="title"> 
          <td bgcolor="#A4B6D7" height="25">|&nbsp; 
            <%
dim sqlBigClass,sqlSmallClass,rsBigClass,rsSmallClass,sqlSpecial,rsSpecial
sqlBigClass="select * from BigClass_down"
Set rsBigClass= Server.CreateObject("ADODB.Recordset")
rsBigClass.open sqlBigClass,conn,1,1
if rsBigClass.eof then 
	response.Write("��û���κ���Ŀ��������������Ŀ��")
end if
do while not rsBigClass.eof
	if rsBigClass("BigClassName")=BigClassName then
		response.Write("<a href='Down_Manage.asp?BigClassName=" & rsBigClass("BigClassName") & "'><font color='red'>" & rsBigClass("BigClassName") & "</font></a> | ")		
	else
		response.Write("<a href='Down_Manage.asp?BigClassName=" & rsBigClass("BigClassName") & "'>" & rsBigClass("BigClassName") & "</a> | ")
	end if
	rsBigClass.movenext
loop
rsBigClass.close
set rsBigClass=nothing
%>
          </td>
        </tr>
        <%
if BigClassName<>"" then
	sqlSmallClass="select * from SmallClass_down where BigClassName='" & BigClassName & "'"
	Set rsSmallClass= Server.CreateObject("ADODB.Recordset")
	rsSmallClass.open sqlSmallClass,conn,1,1
	if not (rsSmallClass.bof and rsSmallClass.eof) then
		response.write "<tr class='tdbg'><td bgcolor='#C0C0C0'>"
		do while not rsSmallClass.eof
			if rsSmallClass("SmallClassName")=SmallClassName then
				response.Write("&nbsp;<a href='Down_Manage.asp?BigClassName=" & rsSmallClass("BigClassName") & "&SmallClassName=" & rsSmallClass("SmallClassName") & "'><font color='red'>" & rsSmallClass("SmallClassName") & "</font></a>&nbsp;&nbsp;")				
			else
				response.Write("&nbsp;<a href='Down_Manage.asp?BigClassName=" & rsSmallClass("BigClassName") & "&SmallClassName=" & rsSmallClass("SmallClassName") & "'>" & rsSmallClass("SmallClassName") & "</a>&nbsp;&nbsp;")
			end if
			rsSmallClass.movenext
		loop
		response.write "</td></tr>"
	end if
	rsSmallClass.close
	set rsSmallClass=nothing
end if
%>
      </table>
      <form name="del" method="Post" action="Down_Del.asp" onsubmit="return ConfirmDel();">
        <table width="620" border="0" cellpadding="0" cellspacing="1" bgcolor="#000000">
          <tr> 
            <td height="25" bgcolor="#A4B6D7"><a href="Down_Manage.asp">&nbsp;������Ѷ����</a> 
              &gt;&gt; 
              <%
if request.querystring="" then
	response.write "��������"
else
	if request("Query")<>"" then
		if Title<>"" then
			response.write "�����к��С�<font color=blue>" & Title & "</font>���Ĳ�Ʒ"
		else
			response.Write("��������")
		end if
 	else
		if BigClassName<>"" then
			response.write "<a href='Down_Manage.asp?BigClassName=" & BigClassName & "'>" & BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if SmallClassName<>"" then
				response.write "<a href='Down_Manage.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "'>" & SmallClassName & "</a>"
			else
				response.write "����С��"
			end if
		end if		
	end if
end if
%>
            </td>
            <td width="150" bgcolor="#A4B6D7">&nbsp; 
              <%
  	if rs.eof and rs.bof then
		response.write "���ҵ� 0 ������</td></tr></table>"
	else
    	totalPut=rs.recordcount
		if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if

    	end if
		response.Write "���ҵ� " & totalPut & " ������"
%>
            </td>
          </tr>
        </table>
        <%		
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,false,"������"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,false,"������"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,false,"������"
	    	end if
		end if
	end if
%>
        <%  
sub showContent
   	dim i
    i=0
%>
        <table width="620" border="0" cellpadding="0" cellspacing="1" bgcolor="#000000" class="border" style="word-break:break-all">
          <tr bgcolor="#A4B6D7" class="title"> 
            <td width="36" height="25" align="center">ѡ��</td>
            <td width="40"  height="25" align="center">ID</td>
            <td width="246" align="center" >���ر���</td>
            <td width="81" align="center" >����һ������</td>
            <td width="76" align="center" >������������</td>
            <td width="57" align="center" >����ʱ��</td>
            <td width="68" align="center" >����</td>
          </tr>
          <%do while not rs.eof%>
          <tr class="tdbg"> 
            <td width="36" height="22" align="center" bgcolor="#A4B6D7"> 
              <input name='ID' type='checkbox' onclick="unselectall()" id="ID" value='<%=cstr(rs("ID"))%>'>
            </td>
            <td width="40" align="center" bgcolor="#ECF5FF"><%=rs("id")%></td>
            <td  bgcolor="#ECF5FF"><a href="../DownloadShow.asp?ID=<%=rs("id")%>" target="_blank"><%=left(rs("title"),18)%></a></td>
            <td  bgcolor="#ECF5FF">
<div align="center"><%=rs("BigClassName")%></div></td>
            <td  bgcolor="#ECF5FF">
<div align="center"><%=rs("SmallClassName")%></div></td>
            <td align="center" bgcolor="#ECF5FF"><%= FormatDateTime(rs("AddDate"),2) %> </td>
            <td width="68" align="center" bgcolor="#ECF5FF"> <a href="Down_modi.asp?ID=<%=rs("id")%>">�޸�</a> 
              <a href="Down_del.asp?ID=<%=rs("ID")%>&Action=Del" onclick="return ConfirmDel();">ɾ��</a> 
            </td>
          </tr>
          <%
	i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
	loop
%>
        </table>
        <table width="620" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="250" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              ѡ�б�ҳ��ʾ����������</td>
            <td><input name="submit" type='submit' value='ɾ��ѡ��������'> 
              <input name="Action" type="hidden" id="Action" value="Del"></td>
          </tr>
        </table>
        <%
   end sub 
%>
      </form>
      <br> <table width="620" border="0" cellpadding="0" cellspacing="0" class="border">
        <tr class="tdbg"> 
          <form name="searchsoft" method="get" action="Down_Manage.asp">
            <td height="30"> <strong>�������أ�</strong> <input name="Title" type="text" class=smallInput id="Title" size="20" maxlength="50"> 
              <input name="Query" type="submit" id="Query" value="�� ѯ"> &nbsp;&nbsp;�������������ơ����Ϊ�գ�������������ء�</td>
          </form>
        </tr>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
rs.close
set rs=nothing  
call CloseConn()
%>
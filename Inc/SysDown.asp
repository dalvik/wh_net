<META http-equiv=Content-Type content="text/html; charset=gb2312"> 
<!--#Include File="Check_Sql.asp"-->
<!--#include file="Conn.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="Function.asp"-->
<%
dim strFileName,MaxPerPage,ShowSmallClassType
dim totalPut,CurrentPage,TotalPages
dim BeginTime,EndTime
dim founderr, errmsg
dim BigClassName,SmallClassName,keyword,strField
dim rs,sql,sqlDown,rsDown,sqlSearch,rsSearch,sqlBigClass,rsBigClass,sqlBigClass_Down
BeginTime=Timer
BigClassName=Trim(request("BigClassName"))
SmallClassName=Trim(request("SmallClassName"))
keyword=trim(request("keyword"))
if keyword<>"" then 
  keyword=replace(replace(replace(replace(keyword,"'","��"),"<","&lt;"),">","&gt;")," ","&nbsp;")
end if
strField=trim(request("Field"))

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

sqlBigClass_Down="select * from BigClass_Down order by BigClassID"
Set rsBigClass_Down= Server.CreateObject("ADODB.Recordset")
rsBigClass_Down.open sqlBigClass_Down,conn,1,1
%>

<%
'=================================================
'��������ShowSmallClass_Down_Tree
'��  �ã�����Ŀ¼��ʽ��ʾ��Ŀ
'��  ������
'=================================================
sub ShowSmallClass_Down_Tree()
%>
<SCRIPT language=javascript>
function opencat(cat,img){
	if(cat.style.display=="none"){
	cat.style.display="";
	img.src="img/class2.gif";
	} else{
	cat.style.display="none"; 
	img.src="img/class1.gif";
	}
}
</Script>
<TABLE cellSpacing=0 cellPadding=0 width="99%" border=0>
<%
dim i
set rsbigdown = server.CreateObject ("adodb.recordset")
	sql="select * from BigClass_down"
	rsbigdown.open sql,conn,1,1
if rsbigdown.eof and rsbigdown.bof then
	Response.Write "��Ŀ���ڽ����С���"
else
	i=1
	do while not rsbigdown.eof
%>
	<TR>
		<TD language=javascript onmouseup="opencat(cat10<%=i%>000,&#13;&#10; img10<%=i%>000);" id=item$pval[catID]) style="CURSOR: hand" width=34 height=24 align=center><IMG id=img10<%=i%>000 src="img/class1.gif" width=15 height=17></TD>
		<TD width="662">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='download.asp?BigClassName=<%=rsbigdown("BigClassName")%>'><%=rsbigdown("BigClassName")%></a></TD>
	</TR>
	<TR>
		<TD id=cat10<%=i%>000 <%if rsbigdown("BigClassName")=BigClassName then 
		     response.write "style='DISPLAY'"   
		    else  
		     response.write "style='DISPLAY: none'"
		    end if%> colspan="2">
<%
dim rsSmall,sqls,j
set rsSmall = server.CreateObject ("adodb.recordset")
sqls="select * from SmallClass_down where BigClassName='" & rsbigdown("BigClassName") & "' order by SmallClassID"
rsSmall.open sqls,conn,1,1
if rsSmall.eof and rsSmall.bof then
	Response.Write "û��С��"
else
	j=1
	do while not rsSmall.eof
%>
&nbsp;<IMG height=20 src="img/class3.gif" width=26 align=absMiddle border=0><a href="download.asp?BigClassName=<%=rsSmall("BigClassName")%>&Smallclassname=<%=rsSmall("SmallClassName")%>"><%=rsSmall("SmallClassName")%></a><BR>
<%
	rsSmall.movenext
	j=j+1
	loop
end if
rsSmall.close
set rsSmall=nothing
%>
		</TD>
	</TR>
<%
	rsbigdown.movenext
	i=i+1
	loop
	rsbigdown.close
    set rsbigdown=nothing
end if
%>
</TABLE>
<%
end sub
%>

<%
'=================================================
'��������ShowClass_DownGuide
'��  �ã���ʾ��Ŀ����λ��
'��  ������
'=================================================
sub ShowClass_DownGuide()
	response.write  "&nbsp;<a href='download.asp'>����</a>&nbsp;&gt;&gt;"
	if BigClassName="" and SmallClassName="" then
		response.write "&nbsp;��������"
	else
		if BigClassName<>"" then
			response.write "&nbsp;<a href='Download.asp?BigClassName=" & BigClassName & "'>" & BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if SmallClassName<>"" then
				response.write "<a href='Download.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "'>" & SmallClassName & "</a>"
			else
				response.write "����С��"
			end if
		end if
	end if
end sub

'=================================================
'��������ShowDownTotal
'��  �ã���ʾ��������
'��  ������
'=================================================
sub ShowDownTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from download "
	if BigClassName<>"" then
		sqlTotal=sqlTotal & " where BigClassName='" & BigClassName & "' "
		if SmallClassName<>"" then
			sqlTotal=sqlTotal & " and SmallClassName='" & SmallClassName & "' "
		end if	
	end if
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "���� 0 ������"
	else
		totalPut=rsTotal(0)
		response.Write "���� " & totalPut & " ������"
	end if
	rsTotal.close
	set rsTotal=nothing
end sub

'=================================================
'��������ShowDown
'=================================================
sub ShowDown(TitleLen)
	if TitleLen<0 or TitleLen>200 then
		TitleLen=50
	end if
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
	if currentPage=1 then
        sqlDown="select top " & MaxPerPage	
	else
		sqlDown="select "
	end if

	sqlDown=sqlDown & " ID,title,content,BigClassName,SmallClassName,System,Language,Softclass,PhotoUrl,DownloadUrl,FileSize,Hits,AddDate from download"
	
	if BigClassName<>"" then
		sqlDown=sqlDown & " where BigClassName='" & BigClassName & "' "
		if SmallClassName<>"" then
			sqlDown=sqlDown & " and SmallClassName='" & SmallClassName & "' "
		end if	
	end if
	sqlDown=sqlDown & " order by AddDate desc"
	Set rsDown= Server.CreateObject("ADODB.Recordset")
	rsDown.open sqlDown,conn,1,1
	if rsDown.bof and  rsDown.eof then
		response.Write("<br><li>û���κ�����</li>")
	else
		if currentPage=1 then
			call DownContent(TitleLen)
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsDown.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsDown.bookmark
            	call DownContent(TitleLen)
        	else
	        	currentPage=1
           		call DownContent(TitleLen)
	    	end if
		end if
	end if
	rsDown.close
	set rsDown=nothing
end sub

sub DownContent(intTitleLen)
   	dim i,strTemp
    i=0
	do while not rsDown.eof
		strTemp=""
		strTemp=strTemp & "<table width=98% border=0 cellSpacing=1 cellpadding=3 bgColor=#b8b8b8>"
			strTemp=strTemp & "<tr><td width=60% height=30 bgColor=#F0F0F0>"
			strTemp=strTemp & "<a href=DownloadShow.asp?ID=" & rsDown("ID") & ">&nbsp;<b>" & rsDown("Title") & "</b></a>"
			strTemp=strTemp & "</td>"
			strTemp=strTemp & "<td bgColor=#F0F0F0 align=right>"
			strTemp=strTemp & "<a href=Download.asp?BigClassName=" & rsDown("BigClassName") & ">" & rsDown("BigClassName") & "</a> | "
			strTemp=strTemp & "<a href=Download.asp?BigClassName=" & rsDown("BigClassName") & "&SmallClassName=" & rsDown("SmallClassName") & ">"
			strTemp=strTemp & rsDown("SmallClassName") & ""
			strTemp=strTemp & "</a>&nbsp;</td></tr>"

			strTemp=strTemp & "<tr><td height=30 bgColor=#ffffff>&nbsp"
			strTemp=strTemp & "�������ڣ�" & FormatDateTime(rsDown("AddDate"),2) & "&nbsp;&nbsp;"
			strTemp=strTemp & "�ļ���С��" & rsDown("FileSize") & "</td>"
			strTemp=strTemp & "<td height=30 bgColor=#ffffff>"
			strTemp=strTemp & "<table width=100% border=0 cellpadding=0 cellspacing=0 bgColor=#ffffff><tr>"
			strTemp=strTemp & "<td width=40% height=22>&nbsp"
			strTemp=strTemp & "<a href=" & rsDown("DownloadUrl") & " target=_blank>����</a> | "
			strTemp=strTemp & "<a href=DownloadShow.asp?ID=" & rsDown("ID") & ">�鿴</a>&nbsp;"
			strTemp=strTemp & "</td><td align=center>"
			strTemp=strTemp & "�������" & rsDown("Hits") & ""
			strTemp=strTemp & "</td></tr></table>"
			strTemp=strTemp & "</td></tr>"	

			strTemp=strTemp & "<tr><td width=60% height=30 bgColor=#ffffff colspan=2>&nbsp;"
			strTemp=strTemp & "�������ԣ�" & rsDown("Language") & "&nbsp;&nbsp;"
			strTemp=strTemp & "���л�����" & rsDown("System") & ""
			strTemp=strTemp & "</td></tr>"
			strTemp=strTemp & ""

			strTemp=strTemp & "<tr><td height=1 colspan=2 bgcolor=#CCCCCC></td>"
			strTemp=strTemp & "</tr></table><BR>"
		response.write strTemp
		rsDown.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop
end sub 

'=================================================
'��������ShowUserLogin
'��  �ã���ʾ�û���¼����
'��  ������
'=================================================
sub ShowUserLogin()
	dim strLogin
	If Session("UserName")="" Then
    	strLogin= "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
		strLogin=strLogin & "<form action='UserLogin.asp' method='post' name='UserLogin' onSubmit='return CheckForm();'>"
        strLogin=strLogin & "<tr><td height='25' align='right'>�û�����</td><td height='25'><input name='UserName' type='text' id='UserName' size='10' maxlength='20'></td></tr>"
        strLogin=strLogin & "<tr><td height='25' align='right'>��&nbsp;&nbsp;�룺</td><td height='25'><input name='Password' type='password' id='Password' size='10' maxlength='20'></td></tr>"
        strLogin=strLogin & "<tr align='center'><td height='25' colspan='2'><input name='Login' type='submit' id='Login' value=' ��¼ '> <input name='Reset' type='reset' id='Reset' value=' ��� '>"
        strLogin=strLogin & "</td></tr>"
        strLogin=strLogin & "<tr><td height='20' align='center' colspan='2'><a href='UserReg.asp' target='_blank'>���û�ע��</a>&nbsp;&nbsp;<a href='GetPassword.asp' target='_blank'>�������룿</a></td></tr>"      
        strLogin=strLogin & "</form></table>"
		response.write strLogin
%>
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
	Else 
		response.write "<br>��ӭ����" & Session("UserName") & "<br><br>"
		response.write "<a href='userServer.asp'><b><font color=ff3300>��Ա��������</font></b></a><br><br>"
	end if
end sub
%>
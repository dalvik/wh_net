<%@language=vbscript codepage=936 %>
<%
response.buffer=true
%>
<!--#include file="conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="Admin.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim sql,rs,Action,FoundErr,ErrMsg
dim JMObjInstalled
Action=trim(request("Action"))
Email=trim(request("Email"))
mane=trim(request("Name"))
JMObjInstalled=IsObjInstalled("JMail.Message")

dim FSObjInstalled
FSObjInstalled=IsObjInstalled("Scripting.FileSystemObject")


%>
<!-- #include file="Inc/Head.asp" -->
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#000000" class="border">
  <tr class="topbg"> 
    <td class="back_southidc" height="28" colspan="2" align="center" bgcolor="#FFFFFF"><strong>�� 
      �� �� �� �� ��</strong></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="tdbg"> 
    <td width="101" height="30" bgcolor="#A4B6D7"><div align="right">����������</div></td>
    <td width="595" height="30"><a href="AdminMaillist.asp">�����ʼ��б�</a> | <a href="AdminMaillist.asp?Action=Export">�����ʼ��б�</a> 
  </tr>
</table>
<br>
<%
if Action="Send" then
	call SendMaillist()
elseif Action="Export" then
	call ExportMail()
elseif Action="DoExport" then
	call DoExportList()
else
	call main()
end if

if FoundErr=True then
	call WriteErrMsg()
end if

sub main()
%>
<form method="POST" action="AdminMaillist.asp?Action=Send">
  <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" Class="border">
    <tr bgcolor="#FFFFFF" class="title"> 
      <td class="back_southidc" height="28" colspan=2 align=center><b> �� �� �� ��</b></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td rowspan="3" align="right" bgcolor="#A4B6D7">�ռ��ˣ�</td>
      <td width="85%"> 
        <input type="radio" name="incepttype" value="1">
        �ʼ���������ע���û�</td>
    </tr>
    <tr class="tdbg"> 
      <td width="85%" bgcolor="#FFFFFF"> 
        <input type="radio" name="incepttype" value="2">
        ���û����������ʼ�&nbsp;
        <input type="text" name="inceptname" size="35">
        ����û�������<font color="#0000FF">Ӣ�ĵĶ���</font>�ָ���</td>
    </tr>
    <tr class="tdbg"> 
      <td width="85%" bgcolor="#FFFFFF"> 
        <input name="incepttype" type="radio" value="3" checked>
        ���û�Email�����ʼ� 		
        <input name="inceptemail" type="text" value="<%=Email%>" size="35">
        ����û�Email����<font color="#0000FF">Ӣ�ĵĶ���</font>�ָ���</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="15%" align="right" bgcolor="#A4B6D7">�ʼ����⣺</td>
      <td width="85%"> 
        <input type=text name=subject size=64>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td align="right" bgcolor="#A4B6D7">�ʼ����ݣ�</td>
      <td> 
        <textarea cols=80 rows=8 name="content"></textarea>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="15%" align="right" bgcolor="#A4B6D7">�����ˣ�</td>
      <td width="85%"> 
        <input type="text" name="sendername" size="64" value="<%=SiteName%>">
        </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="15%" align="right" bgcolor="#A4B6D7">������Email��</td>
      <td width="85%"> 
        <input type="text" name="senderemail" size="64" value="<%=WebMasterEmail%>">
        </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td align="right" bgcolor="#A4B6D7">�ʼ����ȼ���</td>
      <td> 
        <input type="radio" name="Priority" value="1">
        �� 
        <input type="radio" name="Priority" value="3" checked>
        ��ͨ 
        <input type="radio" name="Priority" value="5">
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="15%" align="right" bgcolor="#A4B6D7">ע�����</td>
      <td width="85%"> 
        <%
		If JMObjInstalled=false Then
			Response.Write "<b><font color=red>�Բ�����Ϊ��������֧�� JMail���! ���Բ���ʹ�ñ����ܡ�</font></b>"
		else
			Response.Write "��Ϣ�����͵�����ע��ʱ������д��������û����ʼ��б���ʹ�ý����Ĵ����ķ�������Դ��������ʹ�á�"
		End If
		%>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td colspan=2 align=center> 
        <input name="Action" type="hidden" id="Action" value="Send">
        <input name="Submit" type="submit" id="Submit" value=" �� �� " <% If JMObjInstalled=false Then response.write "disabled" end if%>>
        &nbsp; 
        <input  name="Reset" type="reset" id="Reset2" value=" �� �� ">
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub SendMaillist()
	dim Sendername,Senderemail,Subject,Content,Priority,InceptType,InceptName,InceptEmail,i,j
	Sendername=trim(request("sendername"))
	Senderemail=trim(request("senderemail"))
	Subject=trim(request("Subject"))
	Content=trim(request("Content"))
	Priority=trim(request("Priority"))
	if Sendername="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����˲���Ϊ�գ�</li>"
	end if
	if Senderemail="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>������Email����Ϊ�գ�</li>"
	end if
	if Subject="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�ʼ����ⲻ��Ϊ�գ�</li>"
	end if
	if Content="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�ʼ����ݲ���Ϊ�գ�</li>"
	end if
	if Priority="" then
		Priority=3
	end if
	
	InceptType=Clng(request("incepttype"))
	sql="select UserName,Email from [User] " 
	if InceptType=1 then		
		sql=sql & " where Email like '%@%'"		
	elseif InceptType=2 then
		InceptName=replace(replace(replace(replace(request("inceptname")," ",""),"'",""),chr(34),""),"|","','")
		sql=sql & " where UserName in ('" & InceptName & "') and Email like '%@%'"
	elseif InceptType=3 then
		InceptEmail=replace(replace(replace(replace(request("inceptemail")," ",""),"'",""),chr(34),""),"|","','")
		sql=sql & " where Email in ('" & InceptEmail & "')"
	end if
	
	
	if FoundErr=True then
		exit sub
	end if

	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
	  if InceptType=3  then
	       if IsValidEmail(Email)=True then			   
				ErrMsg=SendMail(Email,mane,Subject,Content,Sendername,Senderemail,Priority)
				if ErrMsg<>"" then
					FoundErr=True					
				end if
				i=1			
		    end if
	  else	  
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ʱû���û�ע�ᣡ</li>"
	  end if	
	else	    
		response.write "<li>���ڷ����У���ȴ� "
		do while not rs.eof		
			if IsValidEmail(rs("Email"))=True then			   
				ErrMsg=SendMail(rs("Email"),rs("UserName"),Subject,Content,Sendername,Senderemail,Priority)
				if ErrMsg<>"" then
					FoundErr=True
					exit sub
				end if
				i=i+1
				response.write "."
			else
				j=j+1				
			end if
			rs.movenext			
		loop
		response.write "<BR><li>�ɹ������ʼ���"&i&"��"
		if j>0 then response.write "<BR><li>δ�����ʼ���"&j&"�⣨�ʼ���ַ���󣩡�" end if
	end if
	rs.close
	set rs=nothing
	call CloseConn()
end sub

sub ExportMail()
%>
<form method="POST" action="AdminMaillist.asp?Action=DoExport">
  <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" Class="border">
    <tr bgcolor="#FFFFFF" class="title"> 
      <td  height="28" colspan=2 align=center class="back_southidc"><b> �ʼ��б��������������ݿ�</b></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="24%" height="80" align="right" bgcolor="#A4B6D7">�����ʼ��б������ݿ⣺</td>
      <td width="76%" height="80"> 
        <input name="ExportType" type="hidden" id="ExportType" value="1">
        &nbsp;&nbsp;<font color=blue>����</font>&nbsp;&nbsp; 
        <select name="UserType" id="UserType">
          <option value="0" selected>ȫ���û�</option>          
        </select>
        &nbsp;<font color=blue>��</font>&nbsp;
        <input name="ExportFileName" type="text" id="ExportFileName" value="Emaillist.mdb" size="30" maxlength="200">
        <input name="Submit1" type="submit" id="Submit1" value="��ʼ">
      </td>
    </tr>
  </table>
</form>
<form method="POST" action="AdminMaillist.asp?Action=DoExport">
  <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" Class="border">
    <tr bgcolor="#FFFFFF" class="title"> 
      <td height="28" colspan=2 align=center class="back_southidc"><b>�ʼ��б������������ı�</b></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="24%" height="80" align="right" bgcolor="#A4B6D7">�����ʼ��б������ݿ⣺</td>
      <td width="76%" height="80"> 
        <input name="ExportType" type="hidden" id="ExportType" value="2">
        &nbsp;&nbsp;<font color=blue>����</font>&nbsp;&nbsp; 
        <select name="UserType" id="UserType">
          <option value="0" selected>ȫ���û�</option>        
        </select>
        &nbsp;<font color=blue>��</font>&nbsp;
		<input name="ExportFileName" type="text" id="ExportFileName" value="Emaillist.txt" size="30" maxlength="200">
        <input type="submit" name="Submit2" value="��ʼ" <%If FSObjInstalled=false Then response.Write "disabled"%>> 
		<%
			If FSObjInstalled=false Then
				Response.Write "<font color=red>��ķ�������֧�� FSO! ����ʹ�ô˹��ܡ�</font>"
			end if
		%>
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub DoExportList()
   
	dim ExportType,UserType,ExportFileName,strResult,i
	ExportType=Clng(trim(Request("ExportType")))	
	ExportFileName=trim(request("ExportFileName"))
	if ExportFileName="" then
		FoundErr=True
		if ExportType=1 then
			ErrMsg=ErrMsg & "<br><li>������Ҫ���������ݿ��ļ�����</li>"
		else
			ErrMsg=ErrMsg & "<br><li>������Ҫ�������ı��ļ�����</li>"
		end if
	else
		ExportFileName=replace(replace(ExportFileName,"'",""),chr(34),"")
	end if	
	
	set rs=server.createobject("adodb.recordset")	
	sql="select Email from [User] where Email like '%@%'"		
	rs.open sql,conn,1,1	
	i=0
	select case ExportType
	case 1
	    	    	
		dim tconn,tconnstr	
		Set tconn = Server.CreateObject("ADODB.Connection")		
		tconnstr="DBQ="+server.mappath(""&ExportFileName&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
		tconn.Open tconnstr
		do while not rs.eof
			tconn.execute("insert into [user] (email) values ('"&rs(0)&"')")			
			rs.movenext
			i=i+1
		loop
		tConn.close
		Set tconn = Nothing
		strResult="�����ɹ��������� "& i &" ���û�Email��ַ�����ݿ� "&tdb&"��<a href="&ExportFileName&">������ｫ���ݿ����ػر���</a>"
	case 2
		dim fso,filepath,writefile
	
		Set fso = CreateObject("Scripting.FileSystemObject")
		Application.lock
		filepath=Server.MapPath(""&ExportFileName&"")
		Set Writefile = fso.CreateTextFile(filepath,true)
		do while not rs.eof
			Writefile.WriteLine rs(0)
			rs.movenext
			i=i+1
		loop
		Writefile.close
		Application.unlock
		set fso=nothing
		strResult="�����ɹ��������� " & i & " ���û�Email��ַ��"&ExportFileName&"�ļ���<a href="&ttxt&">������ｫ�ļ����ػر���</a>)"
	end select
	rs.close
	set rs=nothing
%>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" class="border">
  <tr class="title"> 
    <td height="28" align=center bgcolor="#FFFFFF" class="back_southidc"><b>�ʼ��б���������������Ϣ</b></td>
  </tr>
  <tr class="tdbg"> 
    <td height="100" align="center" bgcolor="#FFFFFF">
<%response.write strResult%></td>
  </tr>
</table>
<%
end sub
%>
</body>
</html>
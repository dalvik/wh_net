<%@language=VBScript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="../Inc/Config.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%


dim rs,sql,ErrMsg,FoundErr
dim ID,Product_Id,BigClassName,EnBigClassName,SmallClassName,EnSmallClassName 
dim Title,EnTitle,Spec,EnSpec,Unit,EnUnit,Memo,EnMemo,Content,EnContent
dim UpdateTime,Hits,DefaultPicUrl,Elite,Passed
ID=Trim(Request.Form("ID"))
Product_Id=trim(request.form("Product_Id"))
BigClassName=trim(request.form("BigClassName"))
SmallClassName=trim(request.form("SmallClassName"))
Title=trim(request.form("Title"))
EnTitle=trim(request.form("EnTitle"))
Spec=trim(request.form("Spec"))
EnSpec=trim(request.form("EnSpec"))
Unit=trim(request.form("Unit"))
EnUnit=trim(request.form("EnUnit"))
Memo=trim(request.form("Memo"))
EnMemo=trim(request.form("EnMemo"))
Content=trim(request.form("Content"))
EnContent=trim(request.form("EnContent"))
UpdateTime=trim(request.form("UpdateTime"))
DefaultPicUrl=trim(request.form("DefaultPicUrl"))
Passed=trim(request.form("Passed"))
Elite=trim(request.form("Elite"))
Newproduct=trim(request.form("Newproduct"))
Hits=trim(request.form("Hits"))

sqlBig="select * from BigClass where BigClassName='" & BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
EnBigClassName=rsBig("EnBigClassName")
rsBig.close
if SmallClassName<>"" then
sqlSmall="select * from SmallClass where SmallClassName='" & SmallClassName & "' and  BigClassName='" & BigClassName & "'"
Set rsSmall= Server.CreateObject("ADODB.Recordset")
rsSmall.open sqlSmall,conn,1,1
EnSmallClassName=rsSmall("EnSmallClassName")
rsSmall.close
end if

if BigClassName="" then
	founderr=true
	errmsg=errmsg+"<li>δָ����Ʒ��������</li>"
end if
if Title="" then
	founderr=true
	errmsg="<li>��Ʒ���ⲻ��Ϊ��</li>"
end if
if EnTitle="" then
	EnTitle=null
end if
if Spec="" then
	Spec=null
end if
if EnSpec="" then
	EnSpec=null
end if
if Unit="" then
	Unit=null
end if
if EnUnit="" then
	EnUnit=null
end if
if Memo="" then
	Memo=null
end if
if EnMemo="" then
	EnMemo=null
end if

if Content="" then
	founderr=true
	errmsg=errmsg+"<li>��Ʒ˵������Ϊ��</li>"
end if
if EnContent="" then
	EnContent=null
end if
che()
if founderr=false then	
	if UpdateTime<>"" and IsDate(UpdateTime)=true then
		UpdateTime=CDate(UpdateTime)
	else
		UpdateTime=now()
	end if
	if Hits<>"" then
		Hits=CLng(Hits)
	else
		Hits=0
	end if
	set rs=server.createobject("adodb.recordset")
	if request("action")="add" then
		sql="select * from Product where (ID is null)"
		rs.open sql,conn,1,3
		rs.addnew
		call SaveData()		
		rs.update
		ID=rs("ID")
		rs.close
		set rs=nothing
	elseif request("action")="Modify" then	  
  		if ID<>"" then
			sql="select * from Product where id=" & ID
			rs.open sql,conn,1,3
			if not (rs.bof and rs.eof) then
				call SaveData()
				rs.update
				rs.close
				set rs=nothing
 			else
				founderr=true
				errmsg=errmsg+"<li>�Ҳ����˲�Ʒ�������Ѿ���ɾ����</li>"
				call WriteErrMsg()
			end if
		else
			founderr=true
			errmsg=errmsg+"<li>����ȷ����ƷID��ֵ</li>"
			call WriteErrMsg()
		end if
	else
		founderr=true
		errmsg=errmsg+"<li>û��ѡ������</li>"
		call WriteErrMsg()
	end if

	call CloseConn()
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="862" align="center" valign="top"> <br> <br> <b> </b><br> <br> <br> 
      <table class="border" align=center width="70%" border="0" cellpadding="0" cellspacing="2" bordercolor="#999999">
        <tr align=center bgcolor="#A4B6D7"> 
          <td  height="25" colspan="2" class="title"><b> 
            <%if request("action")="add" then%>
            ���� 
            <%else%>
            �޸� 
            <%end if%>
            ��Ʒ�ɹ�</b></td>
        </tr>
        <tr> 
          <td width="24%" height="22" bgcolor="#C0C0C0" class="tdbg"> <p align="right">��Ʒ��ţ�</p></td>
          <td width="76%" bgcolor="#E3E3E3" class="tdbg"><%=ID%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">��Ʒ��ţ�</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Product_Id%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">��Ʒ���ƣ�</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Title%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">���</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Spec%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">��λ��</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Unit%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">��ע��</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Memo%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">�������</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"> <%response.write BigClassName
		if SmallClassName<>"" then response.write " &gt;&gt; " & SmallClassName		
		 %> </td>
        </tr>
        <tr> 
          <td height="22" colspan="2" bgcolor="#E3E3E3" class="tdbg"> <p align="center"> 
              ��<a href="ProductModify.asp?ID=<%=ID%>">�޸Ĳ�Ʒ</a>��&nbsp;��<a href="ProductAdd.asp">�������Ӳ�Ʒ</a>��&nbsp;��<a href="ProductManage.asp">��Ʒ����</a>��&nbsp;��<a href="../ProductShow.asp?ID=<%=ID%>" target="_blank">Ԥ����Ʒ����</a>��</p></td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="Inc/Foot.asp" -->
<%
else
	WriteErrMsg
end if

sub SaveData()
	rs("Product_Id")=Product_Id
	rs("BigClassName")=BigClassName
	rs("EnBigClassName")=EnBigClassName
	rs("SmallClassName")=SmallClassName
	rs("EnSmallClassName")=EnSmallClassName
	rs("Title")=Title
	rs("EnTitle")=EnTitle	
	rs("Spec")=Spec
	rs("EnSpec")=EnSpec
	rs("Unit")=Unit
	rs("EnUnit")=EnUnit
	rs("Memo")=Memo
	rs("EnMemo")=EnMemo
	rs("Content")=Content
	rs("EnContent")=EnContent
	rs("Hits")=Hits
	if Passed="yes" then
		rs("Passed")=True
	else
		if EnableProductCheck="No" then
			rs("Passed")=True
		else
			rs("Passed")=False
		end if
	end if
	if Elite="yes" then
		rs("Elite")=True
	else
		rs("Elite")=False
	end if
	if Newproduct="yes" then
		rs("Newproduct")=True
	else
		rs("Newproduct")=False
	end if
	rs("UpdateTime")=UpdateTime	
	rs("DefaultPicUrl")=DefaultPicUrl
end sub	
%>
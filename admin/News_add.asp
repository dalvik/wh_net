<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from SmallClass_New order by SmallClassID asc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("SmallClassName"))%>","<%= trim(rs("BigClassName"))%>","<%= trim(rs("SmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassName.options[document.myform.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    

function AddItem(strFileName){
  document.myform.IncludePic.checked=true;
  document.myform.DefaultPicUrl.value=strFileName;
  document.myform.DefaultPicList.options[document.myform.DefaultPicList.length]=new Option(strFileName,strFileName);
  document.myform.DefaultPicList.selectedIndex+=1;
  if(document.myform.UploadFiles.value==''){
	document.myform.UploadFiles.value=strFileName;
  }
  else{
    document.myform.UploadFiles.value=document.myform.UploadFiles.value+"|"+strFileName;
  }
}

function CheckForm()
{
    if (editor.EditMode.checked==true)
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerText;
    else
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; 

	if (document.myform.title.value.length == 0) {
		alert("���ű���û����д.");
		document.myform.title.focus();
		return false;
	}
	if (document.myform.BigClassName.value.length == 0) {
		alert("���Ŵ���û��ѡ");
		document.myform.BigClassName.focus();
		return false;
	}
	
		if (document.myform.user.value.length == 0) {
		alert("���ŷ�����û����д");
		document.myform.user.focus();
		return false;
	}
	return true;
}
</script>
<!-- #include file="Inc/Head.asp" -->

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top">
	<table width="95%" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <form method="POST" name="myform" onSubmit="return CheckForm();" action="News_save.asp?action=Add" target="_self">
          <tr> 
            <td class="back_southidc" height="30" colspan="2" align="center">��������</td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td width="119" height="25"><font color="#FF0000">*</font>���ű��⣺</td>
            <td width="476"> <input name="title" type="text" class="input" size="50" maxlength="200"></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="25"><font color="#FF0000">*</font>�������</td>
            <td> <%		
        sql = "select * from BigClass_New"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "����������Ŀ��"
		else
		%> <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
			dim selclass
		    selclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
              </select> <select name="SmallClassName">
                <option value="" selected>��ָ��С��</option>
                <%
			sql="select * from SmallClass_New where BigClassName='" & selclass & "'"
			rs.open sql,conn,1,1
			if not(rs.eof and rs.bof) then
			%>
                <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                <% rs.movenext
				do while not rs.eof%>
                <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                <%
			    	rs.movenext
				loop
			end if
	        rs.close
			%>
                <%
			ranNum=int(9*rnd)+10
			iddata=month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
			%>
              </select></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="25" valign="top"><font color="#FF0000">*</font>�������ݣ�</td>
            <td valign="top"> <textarea name="Content" style="display:none"></textarea> 
              <iframe ID="editor" src="../editor.asp" frameborder=1 scrolling=no width="620" height="405"></iframe></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="25">��ҳͼƬ�� 
              <input name="IncludePic" type="hidden" id="IncludePic" value="yes"></td>
            <td><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" size="40" maxlength="120"> 
              <br>
              ��ҳ��ͼƬ,ֱ�Ӵ��ϴ�ͼƬ��ѡ�� 
              <select name="DefaultPicList" id="select" onChange="DefaultPicUrl.value=this.value;">
                <option selected>��ָ����ҳͼƬ</option>
              </select> <input name="UploadFiles" type="hidden" id="UploadFiles"> 
            </td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="25"><font color="#FF0000">*</font>�����ˣ�</td>
            <td> <input name="user" type="text" class="input" value="<%=session("AdminName")%>" size="30"></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="25">��ҳͼƬ���ţ�</td>
            <td> <input type="radio" name="ok" value="Yes">
              �� 
              <input name="ok" type="radio" value="No" checked>
              �� ��<font color="#FF0000">����Ϊ��ҳͼƬ���ţ�ѡ�����ʱ��ע���������Ƿ�������ͼƬ����</font></td>
          </tr>
          <tr> 
            <td height="30" align="center" bgcolor="#ECF5FF"><div align="left">¼��ʱ�䣺</div></td>
            <td height="30" align="center" bgcolor="#ECF5FF"><div align="left"> 
                <input name="AddDate" type="text" id="AddDate" value="<%=now()%>" maxlength="50">
              </div></td>
          </tr>
          <tr> 
            <td height="30" colspan="2" align="center" bgcolor="#ECF5FF"> <input type="submit" name="Submit" value="�ύ" class="input">
              �� 
              <input type="reset" name="Submit" value="����" class="input"> </td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
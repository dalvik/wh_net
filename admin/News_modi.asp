<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim ID,rs_news,FoundErr,ErrMsg
dim sql
dim count
ID=trim(request("ID"))
FoundErr=False
if ID="" then 
	response.Redirect("News_Manage.asp")
end if
sql="select * from News where ID=" & ID & ""
Set rs_news= Server.CreateObject("ADODB.Recordset")
rs_news.open sql,conn,1,1
if FoundErr=True then
	call WriteErrMsg()
else
%>
<%
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
		alert("新闻标题没有填写.");
		document.myform.title.focus();
		return false; 
	}
		if (document.myform.user.value.length == 0) {
		alert("新闻发布人没有填写");
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
      <table  width="95%" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <form method="POST" name="myform" onSubmit="return CheckForm();" action="News_save.asp?action=Modify" target="_self">
          <tr align="center"> 
            <td class="back_southidc" height="30" colspan="2"><font color="#0000FF"><strong>修改新闻</strong></font></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td width="20%" height="24" align="right"><div align="left"><font color="#FF0000">*</font>新 
                闻标题：</div></td>
            <td width="80%" valign="top">　 
              <div align="left">
                <input name="title" type="text" class="input" value="<%=rs_news("title")%>" size="50" maxlength="200">
              </div></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left"><font color="#FF0000">*</font>新闻类别：</div></td>
            <td valign="top"> 　 
              <div align="left">
                <%			
        sql= "select * from BigClass_New"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
		else
		%>
             <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
                  <%
		    do while not rs.eof
			%>
                  <option <% if rs("BigClassName")=rs_news("BigClassName") then response.Write("selected") end if%> value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                  <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
            </select>
                <select name="SmallClassName">
                  <option value="" <%if rs_news("SmallClassName")="" then response.write "selected"%>>不指定小类</option>
                  <%
			sql="select * from SmallClass_New where BigClassName='" & rs_news("BigClassName") & "'" 
			rs.open sql,conn,1,1 
			if not(rs.eof and rs.bof) then  
				do while not rs.eof %>
                  <option <% if rs("SmallClassName")=rs_news("SmallClassName") then response.Write("selected") end if%> value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                  <%
			    	rs.movenext
				loop
			end if
	        rs.close
			%>
                </select>
              </div></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td align="right" valign="top"><div align="left"><font color="#FF0000">*</font>新闻内容：</div></td>
            <td valign="top">
			<textarea name="Content" style="display:none"><%=rs_news("Content")%></textarea> 
              <iframe ID="editor" src="../editor.asp?Action=News&ID=<%=ID%>" frameborder=1 scrolling=no width="620" height="405"></iframe>
		</td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left">首页图片： 
                <input name="IncludePic" type="hidden" id="IncludePic" value="yes">
              </div></td>
            <td valign="top"><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="<%=rs_news("FirstImageName")%>" size="50" maxlength="200">
            </td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left"><font color="#FF0000">*</font>发布人：</div></td>
            <td valign="top">
<div align="left">
                <input name="user" type="text" class="input" size="30" value="<%=rs_news("user")%>">
              </div></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left">首页图片新闻：</div></td>
            <td>　 
              <div align="left">
                <input type="radio" value="Yes" <%if rs_news("Ok")=True then Response.Write "checked"%>  name="OK">
                是 
                <input type="radio" value="No" <%if rs_news("Ok")=False then Response.Write "checked"%> name="Ok">
                否 　<font color="#FF0000">（设为首页图片新闻，选择此项时请注意文章中是否添加有图片！）</font></div></td>
          </tr>
          <tr align="center" bgcolor="#ECF5FF"> 
            <td height="35"><div align="left">录入时间：</div></td>
            <td height="35"><div align="left">
                <input name="AddDate" type="text" id="AddDate" value="<%=rs_news("AddDate")%>" maxlength="50">
              </div></td>
          </tr>
          <tr align="center" bgcolor="#ECF5FF"> 
		   <input type="hidden" name="Id" value="<%=Id%>">
            <td height="35" colspan="2"> <input type="submit" name="Submit" value="提交" class="input">          
              　 
              <input type="reset" name="Submit" value="重置" class="input"> </td>
          </tr>
        </form>
      </table>
  </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
end if
rs_news.close
set rs_news=nothing
call CloseConn()
%>
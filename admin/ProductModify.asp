<%@language=VBScript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim ID,rsProduct,FoundErr,ErrMsg
ID=trim(request("ID"))
if ID="" then 
	response.Redirect("ProductManage.asp")
end if
sql="select * from Product where ID=" & ID & ""
Set rsProduct= Server.CreateObject("ADODB.Recordset")
rsProduct.open sql,conn,1,1
if FoundErr=True then
	call WriteErrMsg()
else
%>
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from SmallClass order by SmallClassID asc"
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

  if (document.myform.Title.value=="")
  {
    alert("产品名称不能为空！");
	document.myform.Title.focus();
	return false;
  }  
  if (document.myform.Content.value=="")
  {
    alert("产品内容不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}

</script>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="862" align="center" valign="top"><br>
      <b> </b>
<form method="POST" name="myform" onSubmit="return CheckForm();" action="ProductSave.asp?action=Modify">
        <table width="650" border="0" align="center" cellpadding="2" cellspacing="1" class="table_southidc">
          <tr> 
            <td class="back_southidc" height="22" colspan="3" align="right" bgcolor="#C0C0C0"><div align="center"><b>修 
                改 产 品</b> </div></td>
          </tr>
          <tr> 
            <td width="167" height="22" align="right" bgcolor="#A4B6D7">所属栏目：</td>
            <td colspan="2" bgcolor="#E3E3E3"> <%
        sql = "select * from BigClass"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
		else
		%> <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
                <%
		    do while not rs.eof
			%>
                <option <% if rs("BigClassName")=rsProduct("BigClassName") then response.Write("selected") end if%> value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
              </select> <%
	%> <select name="SmallClassName">
                <option value="" <%if rsProduct("SmallClassName")="" then response.write "selected"%>>不指定小类</option>
                <%
			sql="select * from SmallClass where BigClassName='" & rsProduct("BigClassName") & "'"
			rs.open sql,conn,1,1
			if not(rs.eof and rs.bof) then
			che()
				do while not rs.eof%>
                <option <% if rs("SmallClassName")=rsProduct("SmallClassName") then response.Write("selected") end if%> value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                <%
			    	rs.movenext
				loop
			end if
	        rs.close
			%>
              </select> </td>
          </tr>
          <tr> 
            <td height="22" align="right" bgcolor="#A4B6D7">产品编号：</td>
            <td colspan="2" bgcolor="#E3E3E3"> <input name="Product_Id" type="text"
           id="Product_Id" value="<%=rsProduct("Product_Id")%>" size="9" maxlength="9"> 
              <font color="#FF0000">*</font> </td>
          </tr>
          <tr> 
            <td height="22" align="right" bgcolor="#A4B6D7">产品名称：</td>
            <td colspan="2" bgcolor="#E3E3E3"> <input name="Title" type="text"
           id="Title" value="<%=rsProduct("Title")%>" size="50" maxlength="80"> 
              <font color="#FF0000">*</font></td>
          </tr>
          <tr> 
            <td height="22" align="right" bgcolor="#A4B6D7">规格：</td>
            <td colspan="2" bgcolor="#E3E3E3"><input name="Spec" type="text"
           id="Spec" value="<%=rsProduct("Spec")%>" size="20" maxlength="80"></td>
          </tr>
          <tr> 
            <td height="22" align="right" bgcolor="#A4B6D7">单位：</td>
            <td colspan="2" bgcolor="#E3E3E3"><input name="Unit" type="text"
           id="spec" value="<%=rsProduct("Unit")%>" size="20" maxlength="80"></td>
          </tr>
          <tr> 
            <td height="22" align="right" bgcolor="#A4B6D7">备注：</td>
            <td colspan="2" bgcolor="#E3E3E3"><input name="Memo" type="text"
           id="unit" value="<%=rsProduct("Memo")%>" size="50" maxlength="80"></td>
          </tr>
          <tr> 
            <td align="right" valign="middle" bgcolor="#A4B6D7"> <div align="right">产品说明： 
              </div></td>
            <td colspan="2" bgcolor="#E3E3E3"><input type="hidden" name="content" value="<%=Server.HTMLEncode(rsProduct("Content"))%>"> 
              <iframe ID="eWebEditor1" src="Southidceditor/ewebeditor.asp?id=content&style=southidc" frameborder="0" scrolling="no" width="550" HEIGHT="450"></iframe>            </td>
          </tr>
          <tr> 
            <td height="22" align="right" bgcolor="#A4B6D7">产品图片：</td>
            <td width="221" bgcolor="#E3E3E3"><input name="DefaultPicUrl" type="text" id="DefaultPicUrl2" value="<%=rsProduct("DefaultPicUrl")%>" size="30" maxlength="50"></td>
            <td width="216" bgcolor="#E3E3E3"><iframe style="top:2px" ID="UploadFiles" src="../upload_Photo.asp?PhotoUrlID=6" frameborder=0 scrolling=no width="300" height="25"></iframe>            </td>
          </tr>
          <tr> 
            <td height="22" align="right" bgcolor="#A4B6D7">已通过审核：</td>
            <td colspan="2" bgcolor="#E3E3E3"> <input name="Passed" type="checkbox" id="Passed" value="yes" <% if rsProduct("Passed")=true then response.Write("checked") end if%>>
              是<font color="#0000FF">（如果选中的话将直接发布）</font></td>
          </tr>
          <tr> 
            <td height="22" align="right" bgcolor="#A4B6D7">首页显示：</td>
            <td colspan="2" bgcolor="#E3E3E3"> <input name="Elite" type="checkbox" id="Elite" value="yes" <% if rsProduct("Elite")=true then response.Write("checked") end if%>>
              是<font color="#0000FF">（如果选中的话将在首页显示）</font></td>
          </tr>
          <tr> 
            <td height="22" align="right" bgcolor="#A4B6D7">首页新品显示：</td>
            <td colspan="2" bgcolor="#E3E3E3"><input name="newproduct" type="checkbox" id="newproduct" value="yes" <% if rsProduct("newproduct")=true then response.Write("checked") end if%>>
              是<font color="#0000FF">（如果选中的话将在首页显示为新品展示）</font></td>
          </tr>
          <tr> 
            <td height="22" align="right" bgcolor="#A4B6D7">录入时间：</td>
            <td colspan="2" bgcolor="#E3E3E3"> <input name="UpdateTime" type="text" id="UpdateTime" value="<%=rsProduct("UpdateTime")%>" maxlength="50">
              当前时间为：<%=date()%> 注意不要改变格式。</td>
          </tr>
          <tr> 
            <td height="22" colspan="3" align="right" bgcolor="#E3E3E3"> <div align="center"> 
                <input  name="Save" type="submit"  id="Save2" value="保存修改结果">
              </div></td>
          </tr>
        </table>
        <div align="center"> 
          <p> 
            <input name="ID" type="hidden" id="ID" value="<%=rsProduct("ID")%>">
          </p>
        </div>
      </form></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
end if
rsProduct.close
set rsProduct=nothing
call CloseConn()
%>
<!--#include file="Inc/SysProduct.asp" -->
<!-- #include file="Head.asp" -->
<SCRIPT language=javaScript>
function CheckJob()
{
	if (document.form1.Quarters.value.length < 2 || document.form1.Quarters.value.length > 30){
		alert ("��ʾ��\n\nְλ������2-30��֮�䣡");
		document.form1.Quarters.focus();
		return false;
	}
	if (document.form1.Name.value.length < 2 || document.form1.Name.value.length > 16){
		alert ("��ʾ��\n\n����������2-16��֮�䣡");
		document.form1.Name.focus();
		return false;
	}
	if (document.form1.Birthday.value.length!=10){
		alert ("��ʾ��\n\n�������ڸ�ʽ���ԣ�");
		document.form1.Birthday.focus();
		return false;
	}
	if (document.form1.Stature.value.length != 3){
		alert ("��ʾ��\n\n��߸�ʽ���ԣ�");
		document.form1.Stature.focus();
		return false;
	}
	if (document.form1.Residence.value.length < 4 ||document.form1.Residence.value.length > 30 ){
		alert ("��ʾ��\n\n�������ڵ���4-30����֮�䣡");
		document.form1.Residence.focus();
		return false;
	}
	if (document.form1.Edulevel.value.length < 20 ){
		alert ("��ʾ��\n\n��������������20�����ϣ�");
		document.form1.Edulevel.focus();
		return false;
	}
	if (document.form1.Experience.value.length < 20 ){
		alert ("��ʾ��\n\n��������������20�����ϣ�");
		document.form1.Experience.focus();
		return false;
	}
	if (document.form1.Phone.value == "" || document.form1.Phone.value.length < 8){
		alert ("��ʾ��\n\n��ϵ�绰������8���ַ����ϣ�");
		document.form1.Phone.focus();
		return false;
	}

    if(document.form1.Email.value.length!=0)
     {
       if (document.form1.Email.value.charAt(0)=="." ||        
            document.form1.Email.value.charAt(0)=="@"||       
            document.form1.Email.value.indexOf('@', 0) == -1 || 
            document.form1.Email.value.indexOf('.', 0) == -1 || 
            document.form1.Email.value.lastIndexOf("@")==document.form1.Email.value.length-1 || 
            document.form1.Email.value.lastIndexOf(".")==document.form1.Email.value.length-1)
        {
         alert("Email��ַ��ʽ����ȷ��");
         document.form1.Email.focus();
         return false;
         }
      }
    else
     {
      alert("Email����Ϊ�գ�");
      document.form1.Email.focus();
      return false;
      }
 }
</SCRIPT>
<body>


<DIV id=index_container>
  <div id="index_left">
    <div id="left">
      <div id="left_title">��ҵ���</div>
      <div id="left_content">
        <DIV id=elem-FrontProductCategory_slideTree-001 name="��Ʒ����"><LINK 
href="imgbyy/css2.css" type=text/css 
rel=stylesheet>

<DIV class=FrontProductCategory_slideTree-default 
id=FrontProductCategory_slideTree-001>
<DIV class=left_bottom>
<DIV class=right_bottom>
<DIV class=left_top>
<DIV class=right_top>


<UL id=treemenu1>

<%
Set rslist = Server.CreateObject("ADODB.Recordset")
sql="select Title from Aboutus where Language='1' order by Aboutusorder"
rslist.open sql,conn,1,3
do while not rslist.eof
%>
  <LI>
  <a href="Aboutus.asp?Title=<%=rslist("Title")%>"><%=rslist("Title")%></a>
  
  </li>
  
  
  
  <%rslist.movenext 
loop%>


  </UL>
  
   <table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10"></td>
  </tr>
</table>
  
</DIV>


</DIV></DIV></DIV></DIV>
</DIV>
      </div>
      <div id="left_btm"></div>
    </div>
	<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="6"></td>
  </tr>
</table>

    <div id="index_search">
      <div id="index_search1"></div>
      <div id="index_search2">
        <div class="query">
          <table width="95%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="5"></td>
            </tr>
          </table>
          <% call ShowSearch(1) %>
          <table width="95%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="5"></td>
            </tr>
          </table>
          <table width="95%" height="65" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="30" align="center"><a onFocus="this.blur()" 
href="msnim:chat?contact=59684111@qq.com"><img 
src="imgbyy/msn.png" alt="MSN" width="158" height="50" border="0" /></a></td>
            </tr>
            <tr>
              <td height="5"></td>
            </tr>
            <tr>
              <td height="30"></td>
            </tr>
          </table>
        </div>
      </div>
      <div id="index_search3"></div>
    </div>
    <div id="index_contact">
      <div id="index_contact1"></div>
      <div id="index_contact2">
        <div id="elem-FrontSpecify_show-001" name="��ϵ����˵��ҳ">
          <link 
href="about_files/css3.css" type="text/css" rel="stylesheet" />
          <div class="FrontSpecify_show-default" id="FrontSpecify_show-001">
            <div class="content">
              <div class="style1"><font face="Arial">��ַ��<font face="Arial"><%=Indexadd%><font 
face="Arial"></font><br />
                </font></font>�绰��<%=Indextel%> <br />
                ���棺<%=Indexfax%><font face="Arial"></font> </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
  <DIV id=index_right><!--<div id="right_title">��˾���</div> -->
<DIV id=right_content>
<DIV ><LINK 
href="about_files/FrontCommonContent_showDetailByList.css" type=text/css 
rel=stylesheet>
<DIV class=FrontCommonContent_showDetailByList-default 
id=>
<DIV class=content>
  <DIV id=right_title>��Ƹ��Ϣ</DIV>
<DIV class=text>
  <table width="100%" height="140%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="206" valign="top"><%
function ChangeChr(str) 
   ChangeChr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;") 
end function
%>
          <%    
dim action,Quarters
Quarters=trim(request.QueryString("Quarters"))
action=trim(request.QueryString("action"))
	   
if action="add" then   
   Quarters=trim(request.Form("Quarters"))    
   Uname=trim(request.Form("Name"))
   Sex=trim(request.Form("Sex"))
   Marry=trim(request.Form("Marry"))
   Birthday=trim(request.Form("Birthday"))
   Stature=trim(request.Form("Stature"))
   School=trim(request.Form("School"))
   Studydegree=trim(request.Form("Studydegree"))
   Specialty=trim(request.Form("Specialty"))
   Gradyear=trim(request.Form("Gradyear"))	
   Residence=trim(request.Form("Residence"))
   Edulevel=trim(request.Form("Edulevel"))
   Edulevel=trim(ChangeChr(Edulevel))
   Experience=trim(request.Form("Experience"))
   Experience=trim(ChangeChr(Experience))
   Phone=trim(request.Form("Phone"))
   Mobile=trim(request.Form("Mobile"))
   Email=trim(request.Form("Email"))
   Add=trim(request.Form("Add"))
   Postcode=trim(request.Form("Postcode"))
   
   if School="" then
	School=null
   end if
   if Studydegree="" then
	Studydegree=null
   end if
   if Specialty="" then
	Specialty=null
   end if
   if Gradyear="" then
	Gradyear=null
   end if
   if Mobile="" then
	Mobile=null
   end if
   if Add="" then
	Add=null
   end if
   if Postcode="" then
	Postcode=null
   end if
 '=================================   
   set rs=server.createobject("adodb.recordset")
       sql="select * from HrDemandAccept where (id is null)" 
       rs.open sql,conn,1,3
       rs.addnew
	   rs("Quarters")=Quarters
	   rs("name")=Uname	 
	   rs("Sex")=Sex	 
	   rs("Marry")=Marry
	   rs("Birthday")=Birthday
	   rs("Stature")=Stature	
	   rs("School")=School
	   rs("Studydegree")=Studydegree
	   rs("Specialty")=Specialty
	   rs("Gradyear")=Gradyear   
	   rs("Residence")=Residence	  
	   rs("Edulevel")=Edulevel
	   rs("Experience")=Experience
	   rs("Phone")=Phone
	   rs("Mobile")=Mobile
	   rs("Email")=Email
	   rs("Add")=Add
	   rs("Postcode")=Postcode
       rs("Adddate")=date()
       rs.update
	   rs.close
	   set rs=nothing
response.write"<SCRIPT language=JavaScript>alert('���ļ����ύ�ɹ�������������ǻᾡ��֪ͨ������!');"
response.write"javascript:history.go(-2)</SCRIPT>"       
else
%>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr valign="top" >
              <td  width="80%" height="18"><form action="HrDemandAccept.asp?action=add" method="post" name="form1" id="form1" onSubmit="return CheckJob()">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><div align="center"></div></td>
                    </tr>
                  </table>
                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC" >
                    <tr bgcolor="#FFFFFF">
                      <%Quarters=request("Quarters")%>
                      <td width="116" height="30"><div align="right"><font color="#000066">ְ����λ��&nbsp; </font></div></td>
                      <td width="548" height="25">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                          <input   name="Quarters"  id="jobname"   value="<%=Quarters%>" size="36" />
                        * </td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td width="116"><div align="right"><font color="#000066">�������ϣ�&nbsp;</font></div></td>
                      <td valign="top"><table width="528" border="0" cellpadding="2" cellspacing="1">
                          <tbody>
                            <tr>
                              <td width="17%" height="20" align="right">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����</td>
                              <td width="83%"><input  name="Name" />
                                * </td>
                            </tr>
                            <tr>
                              <td align="right" 
                                height="20">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</td>
                              <td><select id="Sex" name="Sex">
                                  <option value="��" selected="selected">��</option>
                                  <option value="Ů">Ů</option>
                                </select>
                                *</td>
                            </tr>
                            <tr>
                              <td align="right" 
                                height="20">����״����</td>
                              <td><select id="Marry"   name="Marry">
                                  <option value="δ��"  selected="selected">δ��</option>
                                  <option  value="�ѻ�">�ѻ�</option>
                                </select>
                              </td>
                            </tr>
                            <tr>
                              <td align="right" 
                                height="20">�������ڣ�</td>
                              <td><input name="Birthday" />
                                *�������ڣ��磺1978-04-24��</td>
                            </tr>
                            <tr>
                              <td align="right" 
                                height="20">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ߣ�</td>
                              <td><input name="Stature" id="stature" size="15" maxlength="3" />
                                *(cm)���磺�� 178��</td>
                            </tr>
                            <tr>
                              <td align="right" 
                                height="20">��ҵԺУ��</td>
                              <td><input name="School" size="40" /></td>
                            </tr>
                            <tr>
                              <td align="right" 
                                height="20">ѧ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����</td>
                              <td><input name="Studydegree" size="30" maxlength="50" /></td>
                            </tr>
                            <tr>
                              <td align="right" 
                                height="20">ר&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ҵ��</td>
                              <td><input name="Specialty" size="30" maxlength="50" /></td>
                            </tr>
                            <tr>
                              <td align="right" 
                                height="20">��ҵʱ�䣺</td>
                              <td><input name="Gradyear" size="16" /></td>
                            </tr>
                            <tr>
                              <td align="right" 
                                height="20">�� �� �أ�</td>
                              <td><input 
                                name="Residence" id="Residence" 
                                size="40" maxlength="50" />
                                *</td>
                            </tr>
                          </tbody>
                      </table></td>
                    </tr>
                    <tr>
                      <td width="116" bgcolor="#FFFFFF"><div align="right"><font color="#000066">����������&nbsp;</font></div></td>
                      <td align="center" bgcolor="#FFFFFF"><table width="100%" height="188" 
                                border="0" align="center" cellpadding="2" cellspacing="1">
                          <tbody>
                            <tr bgcolor="#FFFFFF">
                              <td width="16%" 
height="21">&nbsp;ѧ��</td>
                              <td width="22%">&nbsp;��ֹʱ��</td>
                              <td width="22%">&nbsp;רҵ����</td>
                              <td width="19%">&nbsp;֤��</td>
                              <td width="21%">&nbsp;ѧУ����</td>
                            </tr>
                            <tr valign="top" bgcolor="#FFFFFF">
                              <td colspan="5" align="center"><textarea id="Edulevel" name="Edulevel" rows="12" cols="100"></textarea>
                                  <br />
                                * �������������Ҫ������ĸ�ʽ�ͷ���ʱ���Ⱥ���д!</td>
                            </tr>
                          </tbody>
                      </table></td>
                    </tr>
                    <tr>
                      <td width="116" bgcolor="#FFFFFF"><div align="right"><font color="#000066">����������&nbsp;</font></div></td>
                      <td bgcolor="#FFFFFF"><table width="100%" height="188" 
                                border="0" align="center" cellpadding="2" cellspacing="1">
                          <tbody>
                            <tr bgcolor="#FFFFFF">
                              <td width="25%" 
height="21">&nbsp;��ֹʱ��</td>
                              <td width="25%">&nbsp;ְλ����</td>
                              <td width="25%">&nbsp;����λ</td>
                              <td width="25%">&nbsp;��������</td>
                            </tr>
                            <tr valign="top" bgcolor="#FFFFFF">
                              <td colspan="4" align="center"><textarea id="Experience" name="Experience" rows="12" cols="100"></textarea>
                                  <br />
                                * �������������Ҫ������ĸ�ʽ�ͷ���ʱ���Ⱥ���д!</td>
                            </tr>
                          </tbody>
                      </table></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td width="116"><div align="right"><font color="#000066">��ϵ��ʽ��&nbsp;</font> </div></td>
                      <td><table width="500" border="0" cellpadding="2" cellspacing="1">
                          <tbody>
                            <tr>
                              <td width="17%" height="20" align="right">��ϵ�绰��</td>
                              <td width="83%"><input name="Phone" id="Phone" size="25" />
                                * </td>
                            </tr>
                            <tr>
                              <td 
                                height="20" align="right">�ֻ����룺</td>
                              <td><input 
                                name="Mobile" id="Mobile" size="25" />
                              </td>
                            </tr>
                            <tr>
                              <td 
                                height="20" align="right">E-mail��ַ��</td>
                              <td><input 
                                name="Email" id="Email" size="25" />
                                *</td>
                            </tr>
                            <tr>
                              <td 
                                height="20" align="right">ͨ�ŵ�ַ��</td>
                              <td><input name="Add" id="Add" 
                               size="40" /></td>
                            </tr>
                            <tr>
                              <td 
                                height="20" align="right">�������룺</td>
                              <td><input 
                                name="Postcode" id="Postcode" size="10" maxlength="6" /></td>
                            </tr>
                          </tbody>
                      </table></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td width="116">&nbsp;</td>
                      <td><div align="center">
                          <input type="submit" name="submit8" value=" �ύ ">
                        &nbsp;&nbsp;
                        <input type="reset" name="Submit8" value=" ���� " />
                      </div></td>
                    </tr>
                  </table>
              </form></td>
            </tr>
          </table>
        <% end if %>
      </td>
    </tr>
  </table>
</DIV>
</DIV></DIV></DIV></DIV>
<DIV class=clear></DIV></DIV>
<DIV class=clear></DIV>
</DIV></DIV>
<DIV class=clear></DIV>

<DIV id=footer>
<DIV id=elem-FrontSpecify_show-020 ><LINK 
href="imgbyy/css3.css" type=text/css rel=stylesheet>
<DIV class=FrontSpecify_show-default id=FrontSpecify_show-020>
<DIV class=content>
<!--#include file="Inc/Foot.asp" -->


</DIV></DIV></DIV></DIV>
</BODY></HTML>

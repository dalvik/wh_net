<!--#include file="Inc/SysProduct.asp" -->
<%
'����Ķ���������д���
MaxPerPage=MaxPerPage_Default
strFileName="Product.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName 
%>
<!-- #include file="Head.asp" -->
<TABLE cellSpacing=0 cellPadding=0 width=1002 align=center border=0>
  <TBODY>
    <TR>
      <TD vAlign=top width=210><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY>
            <TR>
              <TD><IMG height=40 src="imgbly/lt14.gif" width=210></TD>
            </TR>
            <TR>
              <TD align=middle><TABLE width=180 border=0 align="center" cellPadding=0 
            cellSpacing=0 
            style="BORDER-RIGHT: #e9e9e9 1px solid; BORDER-TOP: #e9e9e9 1px solid; BORDER-LEFT: #e9e9e9 1px solid; BORDER-BOTTOM: #e9e9e9 1px solid">
                <TBODY>
                  <TR>
                    <TD><TABLE cellSpacing=0 cellPadding=0 width=180 
                  border=0>
                        <TBODY>
                          <TR>
                            <TD width="22" align="right">&nbsp;</TD>
                            <TD width="158"><% call ShowSmallClass_Tree() %></TD>
                          </TR>
                        </TBODY>
                    </TABLE></TD>
                  </TR>
                </TBODY>
              </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE>
          <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
            <TBODY>
              <TR>
                <TD><IMG height=38 src="imgbly/tit_sel.gif" width=210></TD>
              </TR>
              <TR>
                <TD align=middle><table width="180" border=0 align="center" cellpadding=0 
                        cellspacing=1 
                        style="BORDER-RIGHT: #e3e3e3 1px solid; BORDER-TOP: #e3e3e3 1px solid; BORDER-LEFT: #e3e3e3 1px solid; BORDER-BOTTOM: #e3e3e3 1px solid">
                    <tbody>
                      <tr>
                        <td style="PADDING-BOTTOM: 8px; PADDING-TOP: 8px"><table border="0" cellpadding="2" cellspacing="0" align="center">
                            <form method="Get" name="myform" action="search.asp">
                              <tr>
                                <td height="28"><input type="text" name="keyword"  style="BORDER-RIGHT: #b7b7b7 1px solid; BORDER-TOP: #b7b7b7 1px solid; BORDER-LEFT: #b7b7b7 1px solid; WIDTH: 120px; BORDER-BOTTOM: #b7b7b7 1px solid" size=12 value="�ؼ���" maxlength="50" onFocus="this.select();">
                                    <input type="submit" name="Submit"  value="����">
                                </td>
                              </tr>
                            </form>
                        </table></td>
                      </tr>
                  </table></TD>
              </TR>
            </TABLE>
          <SCRIPT type=text/ecmascript>
function doLogin(fm){
   if(fm.uid.value=='' || fm.uid.value.indexOf("\'")!=-1){
		alert('�Բ����û�������Ϊ�ջ��е����ţ�');
		fm.uid.focus();
		return false;   			
   }
   if(fm.pwd.value==''){
		alert('�Բ������벻��Ϊ�գ�');
		fm.pwd.focus();
		return false;   			
   }

	var postData="Act=doLogin&uid="+escape(fm.uid.value)+"&pwd="+escape(fm.pwd.value);
	var loginAjax=new sc.Ajax("doLogin.asp", "GET", postData);
	loginAjax.addEventListener(loginEvent);
	loginAjax.Exec();

	function loginEvent(){
		var a=loginAjax.ajax;
		if(a.readyState==4){ 
			if(a.status==200){
				var rtn=a.responseText;
				if(rtn==1){
					alert('��ϲ����������¼�ɹ���');
					document.location.reload();
				}else{
					alert('�Բ��𣬵�¼ʧ�ܣ��û���/�������');
				}
			}else{
				alert("��Ǹ��������æ�����Ժ�����..."+a.statusText );
			}
		}
	}
	return false;
}

function doLogout(){
  var loginAjax=new sc.Ajax('doLogin.asp', 'GET', 'Act=doLogout');
  loginAjax.addEventListener(logoutEvent);
  loginAjax.Exec();

	function logoutEvent(){
		var a=loginAjax.ajax;
		if(a.readyState==4){ 
			if(a.status==200){
				var rtn=a.responseText;
				if(rtn==1){
					alert('ȷ��Ҫ�˳���');
					document.location.href=".";
				}else{
					alert('�Բ����˳�ʧ�ܣ�');
				}
			}else{
				alert("��Ǹ��������æ�����Ժ�����..."+a.statusText );
			}
		}
	}
}

  </SCRIPT>
          <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
            <TBODY>
              <TR>
                <TD><IMG height=38 src="imgbly/tit_mem.gif" width=211></TD>
              </TR>
              <TR>
                <TD align=middle><TABLE 
            style="BORDER-RIGHT: #e3e3e3 1px solid; BORDER-TOP: #e3e3e3 1px solid; BORDER-LEFT: #e3e3e3 1px solid; BORDER-BOTTOM: #e3e3e3 1px solid" 
            cellSpacing=1 cellPadding=0 width=178 align=center border=0>
                    <TBODY>
                      <TR>
                        <TD align="center" 
                style="PADDING-RIGHT: 5px; PADDING-LEFT: 2px; PADDING-BOTTOM: 5px; PADDING-TOP: 6px"><% call ShowUserLogin() %></TD>
                      </TR>
                  </TABLE></TD>
              </TR>
            </TBODY>
        </TABLE></TD>
      <TD style="BACKGROUND: url(imgbly/bg_body.gif) repeat-y 50% top" vAlign=top><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY>
            <TR>
              <TD><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                  <TBODY>
                    <TR>
                      <TD background=imgbly/bg_rt.gif><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                          <TBODY>
                            <TR>
                              <TD><IMG src="imgbly/rt17.gif" width="209" height=41></TD>
                              <TD>&nbsp;</TD>
                              <TD class=loc 
                      style="BACKGROUND: url(img/bg_rtjiao.gif) no-repeat right 50%; COLOR: #9f9f9f" 
                      align=right>�����ڵ�λ�ã�<A href="index.asp">��ҳ</A> &gt; <A href="">��Ʒ����</A> &gt; <A 
                        href="">��Ʒ����</A></TD>
                            </TR>
                          </TBODY>
                      </TABLE></TD>
                      <TD width=17>&nbsp;</TD>
                    </TR>
                  </TBODY>
              </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE>
          <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
            <TBODY>
              <TR>
                <TD width=17 height=15>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD width=17>&nbsp;</TD>
              </TR>
              <TR>
                <TD>&nbsp;</TD>
                <TD style="LINE-HEIGHT: 24px"><% call ShowAllClass() %></TD>
                <TD>&nbsp;</TD>
              </TR>
              <TR>
                <TD height=10>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
              </TR>
            </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
<!-- #include file="Inc/Foot.asp" -->
</BODY></HTML>
<%@language=vbscript codepage=936 %>
<%
option explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'主要是使随机出现的图片数字随机
%>
<script language=javascript>
<!--
function SetFocus()
{
if (document.Login.UserName.value=="")
	document.Login.UserName.focus();
else
	document.Login.UserName.select();
}
function CheckForm()
{
	if(document.Login.UserName.value=="")
	{
		alert("请输入用户名！");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("请输入密码！");
		document.Login.Password.focus();
		return false;
	}
	if (document.Login.CheckCode.value==""){
       alert ("请输入您的验证码！");
       document.Login.CheckCode.focus();
       return(false);
    }
}

function CheckBrowser() 
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("友情提示：\n    你使用的是Netscape浏览器，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  } 
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("友情提示：\n    您的浏览器版本太低，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  }
}
//-->
</script>
<html>
<head>
<title>登陆网站</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
INPUT {
	BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 20px
}
TD {
	FONT-SIZE: 12px
}
.inll {
	BORDER-RIGHT: #ffffff; BORDER-TOP: #ffffff; BORDER-LEFT: #ffffff; BORDER-BOTTOM: #ffffff
}
.STYLE1 {color: #FFFFFF}
.STYLE3 {color: #FFFFFF; font-weight: bold; }
</STYLE>
</head>
<body bgcolor="#FFFFFF" background="" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="308" valign="middle"><TABLE cellSpacing=0 cellPadding=0 width=100% align=center 
bgColor=#ffffff border=0>
        <TBODY>
          <TR> 
            <TD height=182 align="center" vAlign=bottom>　</TD>
          </TR>
          
          <TR>
            <TD height=8 valign="bottom"></TD>
          </TR>
          <TR> 
            <TD height="160" align="center" valign="bottom" background="Image/line.gif"> 
              <TABLE width="59%" align=center border=0>
                <TBODY>
                  <TR>
                    <TD width="32%" height="183" align="center" valign="middle"><img src="Image/Login_pic.gif" width="130" height="106"></TD> 
                    <TD width="3%" valign="top"><img src="Image/Login_line.gif" width="2" height="150"></TD>
                    <TD width="65%"><form name="Login" action="../admin/Admin_ChkLogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">
                        <TABLE width="100%" border=0 align=center cellPadding=0 cellSpacing=0>
                          <TBODY>
                            <TR>
                              <TD height=26 colspan="2" noWrap><img src="Image/Login_tit.gif" width="370" height="42"></TD>
                            </TR>
                            <TR> 
                              <TD width="31%" height=26 align="right" noWrap><span class="STYLE3">用户名：</span></TD>
                              <TD width="69%" height=26 noWrap> 
                              <input name="UserName"  type="text"  id="UserName4" maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></TD>
                            </TR>
                            <TR>
                              <TD height=26 align="right" noWrap class="STYLE3">密　码：</TD>
                              <TD noWrap height=26> 
                                <input name="Password"  type="password" maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></TD>
                            </TR>
                            <TR> 
                              <TD height=26 align="right" noWrap class="STYLE3">验证码：</TD>
                              <TD noWrap height=26> 
                                <input name="CheckCode" size="6" maxlength="4" style="border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); ">
				  <img src="../admin/inc/checkcode.asp"></TD>
                            </TR>
                            <TR> 
                              <TD height=26 colSpan=2 align="center"><input   type="submit" name="Submit" value=" 确&nbsp;认 " style="font-size: 9pt; height: 19; width: 60; color: #000000; background-color: #E1F4EE; border: 1 solid #E1F4EE" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#E1F4EE'">
                &nbsp;
                <input name="reset" type="reset"  id="reset" value=" 清&nbsp;除 " style="font-size: 9pt; height: 19; width: 60; color: #000000; background-color: #E1F4EE; border: 1 solid #E1F4EE" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#E1F4EE'"></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                      </form></TD>
                  </TR>
                </TBODY>
            </TABLE></TD>
          </TR>
        </TBODY>
    </TABLE></td>
  </tr>
  <tr>
    <td height="20" align="center"><span style="margin:5px; color: #333333">如果您没有授权，请不要登录，我们将记录您的操作日志。 版权所有：</span><a href="http://www.jiumu.net" style="text-decoration: none"><font color="#333333">九牧互联</font></a></td>
  </tr>
</table>
<script language="JavaScript" type="text/JavaScript">
CheckBrowser();
SetFocus(); 
</script>
</body>
</html>
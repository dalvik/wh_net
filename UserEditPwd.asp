<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim Action,UserName,FoundErr,ErrMsg
dim rsUser,sqlUser
Action=trim(request("Action"))
UserName=trim(request("UserName"))
if Action="" and session("UserName")="" then
	response.redirect "UserServer.asp"
end if
if Action="Modify" and UserName<>"" then
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from [User] where UserName='" & UserName & "'"
	rsUser.Open sqlUser,conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
	else
		dim OldPassword,Password,PwdConfirm
		OldPassword=trim(request("OldPassword"))
		Password=trim(request("Password"))
		PwdConfirm=trim(request("PwdConfirm"))
		if OldPassword="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请输入旧密码！</li>"
		else
			if Instr(OldPassword,"=")>0 or Instr(OldPassword,"%")>0 or Instr(OldPassword,chr(32))>0 or Instr(OldPassword,"?")>0 or Instr(OldPassword,"&")>0 or Instr(OldPassword,";")>0 or Instr(OldPassword,",")>0 or Instr(OldPassword,"'")>0 or Instr(OldPassword,",")>0 or Instr(OldPassword,chr(34))>0 or Instr(OldPassword,chr(9))>0 or Instr(OldPassword,"")>0 or Instr(OldPassword,"$")>0 then
				errmsg=errmsg+"<br><li>旧密码中含有非法字符</li>"
				founderr=true
			else
				if md5(OldPassword)<>rsUser("Password") then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>你输入的旧密码不对，没有权限修改！</li>"
				end if
			end if
		end if
		if strLength(Password)>12 or strLength(Password)<6 then
			founderr=true
			errmsg=errmsg & "<br><li>请输入新密码(不能大于12小于6)。</li>"
		else
			if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
				errmsg=errmsg+"<br><li>新密码中含有非法字符</li>"
				founderr=true
			end if
		end if
		if PwdConfirm="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请输入确认密码！</li>"
		else
			if PwdConfirm<>Password then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>确认密码与新密码不一致！</li>"
			end if
		end if
		if FoundErr<>true then
			rsUser("Password")=md5(Password)
			rsUser.update
		end if
	end if
	rsUser.close
	set rsUser=nothing
	if FoundErr=True then
		call WriteErrMsg()
	else
	    response.write"<SCRIPT language=JavaScript>alert('成功修改密码！');"
        response.write"javascript:history.go(-1)</SCRIPT>"
	end if
else
%>
<!-- #include file="Head.asp" -->
<body>
<DIV id=index_container>
  <div id="index_left">
    <div id="left">
      <div id="left_title">客户服务</div>
      <div id="left_content">
        <DIV id=elem-FrontProductCategory_slideTree-001 name="商品分类"><LINK 
href="imgbyy/css2.css" type=text/css 
rel=stylesheet>

<DIV class=FrontProductCategory_slideTree-default 
id=FrontProductCategory_slideTree-001>
<DIV class=left_bottom>
<DIV class=right_bottom>
<DIV class=left_top>
<DIV class=right_top>


<UL id=treemenu1>


  <LI>
  <a href="UserEdit.asp">修改会员资料</a>
    </li>
  
    <LI>
  <a href="UserEditPwd.asp">修改会员密码</a>
  </li>
     <LI>
  <a href="UserOrder.asp">购物订单查询</a>
  </li>
    <LI>
  <a href="FeedbackMember.asp">站内留言中心</a>
  </li>
    <LI>
  <a href="UserLogout.asp?Action=Ch">退出会员中心</a>
  </li>
 
  


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

    <div id="index_contact">
      <div id="index_contact1"></div>
      <div id="index_contact2">
        <div id="elem-FrontSpecify_show-001" name="联系我们说明页">
          <link 
href="about_files/css3.css" type="text/css" rel="stylesheet" />
          <div class="FrontSpecify_show-default" id="FrontSpecify_show-001">
            <div class="content">
              <div class="style1"><font face="Arial">地址：<font face="Arial"><%=Indexadd%><font 
face="Arial"></font><br />
                </font></font>电话：<%=Indextel%> <br />
                传真：<%=Indexfax%><font face="Arial"></font> </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
  <DIV id=index_right><!--<div id="right_title">公司简介</div> -->
<DIV id=right_content>
<DIV ><LINK 
href="about_files/FrontCommonContent_showDetailByList.css" type=text/css 
rel=stylesheet>
<DIV class=FrontCommonContent_showDetailByList-default 
id=>
<DIV class=content>
  <DIV id=right_title>用户中心</DIV>
<DIV class=text>
  <form action="UserEditPwd.asp" method="post" name="Form1" id="Form1">
    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E3E3E3">
      <tr align="center" bgcolor="#FFFFFF"  >
        <td height="40" colspan="2"><font color="#FFffff" class="font14"><b>修改</b></font><font color="#FF6600" class="font14"><b>修改用户密码</b></font></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="270" height="45" align="right">用 户 名：</td>
        <td width="394"><%=session("UserName")%>
            <input name="UserName" type="hidden" value="<%=session("UserName")%>" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="270" height="45" align="right">旧 密 码：</td>
        <td><input   type="password" maxlength="16" size="30" name="OldPassword" />
        </td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="270" height="45" align="right">新 密 码：</td>
        <td><input   type="password" maxlength="16" size="30" name="Password" />
        </td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="270" height="45" align="right">确认密码：</td>
        <td><input name="PwdConfirm"   type="password" id="PwdConfirm" size="30" maxlength="16" />
        </td>
      </tr>
      <tr align="center" bgcolor="#FFFFFF">
        <td height="40" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify" />
            <input name="Submit"   type="submit" id="Submit" value=" 保 存 " />
        </td>
      </tr>
    </table>
  </form>
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
<%
end if
call CloseConn()
%>
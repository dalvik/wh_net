<%
'*************************************************
'函数名：gotTopic
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'       strlen ----截取长度
'返回值：截取后的字符串
'*************************************************
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "…"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'***********************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'pos=InStr(1,"abcdefg","cd") 
'则pos会返回3表示查找到并且位置为第三个字符开始。
'这就是“查找”的实现，而“查找下一个”功能的
'实现就是把当前位置作为起始位置继续查找。
'***********************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

'***********************************************
'过程名：showpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'***********************************************
sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
  
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></form></table>"
	response.write strTemp	
end sub

'***********************************************
'过程名：enshowpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'***********************************************
sub enshowpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "Total <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "First  Previous&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>First</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>Previous</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "Next  Last"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>Next</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>Last</a>"
  	end if
   	strTemp=strTemp & "&nbsp;Page No.:<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>page "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/page"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;Turn to:<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">No." & i & "page</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></form></table>"
	response.write strTemp
end sub



'********************************************
'函数名：IsValidEmail
'作  用：检查Email地址合法性
'参  数：email ----要检查的Email地址
'返回值：True  ----Email地址合法
'       False ----Email地址不合法
'********************************************
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

'***************************************************
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'       False ----没有安装
'***************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function


'**************************************************
'函数名：strLength
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：str  ----要求长度的字符串
'返回值：字符串长度
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("中国")=2)
	if WINNT_CHINESE then
        dim l,t,c
        dim i
        l=len(str)
        t=l
        for i=1 to l
        	c=asc(mid(str,i,1))
            if c<0 then c=c+65536
            if c>255 then
                t=t+1
            end if
        next
        strLength=t
    else 
        strLength=len(str)
    end if
    if err.number<>0 then err.clear
end function

'****************************************************
'函数名：SendMail
'作  用：用Jmail组件发送邮件
'参  数：ServerAddress  ----服务器地址
'        AddRecipient  ----收信人地址
'        Subject       ----主题
'        Body          ----信件内容
'        Sender        ----发信人地址
'****************************************************
function SendMail(MailServerAddress,AddRecipient,Subject,Body,Sender,MailFrom)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.SMTPMail")
	if err then
		SendMail= "<br><li>没有安装JMail组件</li>"
		err.clear
		exit function
	end if
	JMail.Logging=True
	JMail.Charset="gb2312"
	JMail.ContentType = "text/html"
	JMail.ServerAddress=MailServerAddress
	JMail.AddRecipient=AddRecipient
	JMail.Subject=Subject
	JMail.Body=MailBody
	JMail.Sender=Sender
	JMail.From = MailFrom
	JMail.Priority=1
	JMail.Execute 
	Set JMail=nothing 
	if err then 
		SendMail=err.description
		err.clear
	else
		SendMail="OK"
	end if
end function

'****************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'****************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=2 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td height='20' class='title'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td height='100' class='tdbg' valign='top'><b>产生错误的可能原因：</b><br>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td class='title'><a href='javascript:history.go(-1)'>【返回】</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'****************************************************
'过程名：WriteSuccessMsg
'作  用：显示成功提示信息
'参  数：无
'****************************************************
sub WriteSuccessMsg(SuccessMsg)
	dim strSuccess
	strSuccess=strSuccess & "<html><head><title>成功信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strSuccess=strSuccess & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strSuccess=strSuccess & "<table cellpadding=2 cellspacing=2 border=0 width=400 class='border' align=center>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center'><td height='20' class='title'><strong>恭喜你！</strong></td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr><td height='100' class='tdbg' valign='top'><br>" & SuccessMsg &"</td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center'><td class='title'><a href='javascript:history.go(-1)'>【返回】</a></td></tr>" & vbcrlf
	strSuccess=strSuccess & "</table>" & vbcrlf
	strSuccess=strSuccess & "</body></html>" & vbcrlf
	response.write strSuccess
end sub

function getFileExtName(fileName)
    dim pos
    pos=instrrev(filename,".")
    if pos>0 then 
        getFileExtName=mid(fileName,pos+1)
    else
        getFileExtName=""
    end if
end function 


'==================================================
'过程名：MenuJS
'作  用：生成下拉菜单相关的JS代码
'参  数：无
'==================================================
sub MenuJS()
	response.write "<script type='text/javascript' language='JavaScript1.2' src='Inc/Southidcmenu.js'></script>"
end sub

dim pNum,pNum2
pNum=1
pNum2=0
'=================================================
'过程名：ShowRootClass_Menu
'作  用：显示一级栏目（下拉菜单效果）
'参  数：Language     -----语言    1-中文  2-英文   
'=================================================
sub ShowRootClass_Menu(Language)
	response.write "<script type='text/javascript' language='JavaScript1.2'>" & vbcrlf & "<!--" & vbcrlf
	response.write "stm_bm(['uueoehr',400,'','images/blank.gif',0,'','',0,0,0,0,0,1,0,0]);" & vbcrlf
	response.write "stm_bp('p0',[0,4,0,0,2,2,0,0,100,'',-2,'',-2,90,0,0,'#000000','transparent','',3,0,0,'#000000']);" & vbcrlf
	response.write "stm_ai('p0i0',[0,'','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#000000','#000000','#FFFFFF','','',0,0]);" & vbcrlf
	if Language=1 then
	response.write "stm_aix('p0i1','p0i0',[1,'网站首页','','',-1,-1,0,'index.asp ','_self','index.asp','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#ff0000','#FFFFFF','#FFFFFF','','']);" & vbcrlf
	else
	response.write "stm_aix('p0i1','p0i0',[1,'Home','','',-1,-1,0,'Enindex.asp ','_self','Enindex.asp','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#ff0000','#FFFFFF','#FFFFFF','','']);" & vbcrlf 
	end if 
	response.write "stm_aix('p0i2','p0i0',[0,'·','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#000000','#FFFFFF','#FFFFFF','','',0,0]);" & vbcrlf

	dim sqlRoot,rsRoot,j
	if Language=1 then
	  sqlRoot="select ClassID,ClassName,Depth,NextID,LinkUrl,Child,Readme From MenuClass"
	else
	  sqlRoot="select ClassID,ClassName,Depth,NextID,LinkUrl,Child,Readme From EnMenuClass"
	end if  
	sqlRoot= sqlRoot & " where Depth=0 and ShowOnTop=True order by RootID"
	Set rsRoot= Server.CreateObject("ADODB.Recordset")
	rsRoot.open sqlRoot,conn,1,1
	if not(rsRoot.bof and rsRoot.eof) then 
		j=3
		do while not rsRoot.eof
			if rsRoot(4)<>"" then
				response.write "stm_aix('p0i"&j&"','p0i0',[1,'" & rsRoot(1) & "','','',-1,-1,0,'" & rsRoot(4) & "','_self','" & rsRoot(4) & "','" & rsRoot(6) & "','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#ff0000','#FFFFFF','#cc0000','','']);" & vbcrlf							
			end if
			if rsRoot(5)>0 then
			    if Language=1 then
				  call GetClassMenu(rsRoot(0),0,1)
				else
				  call GetClassMenu(rsRoot(0),0,2) 
				end if   
			end if
			j=j+1
			response.write "stm_aix('p0i2','p0i0',[0,'·','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#000000','#FFFFFF','#FFFFFF','','',0,0]);" & vbcrlf 			
			j=j+1
			rsRoot.movenext
		loop
	end if
	rsRoot.close
	set rsRoot=nothing
	response.write "stm_em();" & vbcrlf
	response.write "//-->" & vbcrlf & "</script>" & vbcrlf	
end sub

sub GetClassMenu(ID,ShowType,Language)
	dim sqlClass,rsClass,k
	'1,4,0,4,2,3,6,7,100前4个数字控制菜单位置和大小
	if pNum=1 then
		response.write "stm_bp('p" & pNum & "',[0,4,0,4,2,3,6,7,100,'progid:DXImageTransform.Microsoft.Fade(overlap=.5,enabled=0,Duration=0.43)',-2,'',-2,67,2,3,'#999999','#EBEBEB','',3,1,1,'#aca899']);" & vbcrlf
	else
		if ShowType=0 then
			response.write "stm_bpx('p" & pNum & "','p" & pNum2 & "',[1,4,0,0,2,3,6]);" & vbcrlf
		else
			response.write "stm_bpx('p" & pNum & "','p" & pNum2 & "',[1,2,-2,-3,2,3,0]);" & vbcrlf
		end if
	end if
	
	k=0
	if Language=1 then
	 sqlClass="select ClassID,ClassName,Depth,NextID,LinkUrl,Child,Readme From MenuClass"
	else
	   sqlClass="select ClassID,ClassName,Depth,NextID,LinkUrl,Child,Readme From EnMenuClass"
	end if   
	sqlClass= sqlClass & " where ParentID=" & ID & " order by OrderID asc"
	Set rsClass= Server.CreateObject("ADODB.Recordset")
	rsClass.open sqlClass,conn,1,1
	do while not rsClass.eof
		if rsClass(4)<>"" then
			if rsClass(5)>0 then
				response.write "stm_aix('p"&pNum&"i"&k&"','p"&pNum2&"i0',[0,'<span class=menu_txt>" & rsClass(1) & "</span>','','',-1,-1,0,'" & rsClass(4) & "','_self','" & rsClass(4) & "','" & rsClass(6) & "','','',6,0,0,'images/arrow_r.gif','images/arrow_w.gif',7,7,0,0,1,'#FFFFFF',0,'#cccccc',0,'','',3,3,0,0,'#fffff7','#f52087','#f52087','#cc0000','']);" & vbcrlf
				pNum=pNum+1
				pNum2=pNum2+1
				if Language=1 then
				  call GetClassMenu(rsClass(0),1,1)
				else
				  call GetClassMenu(rsClass(0),1,2) 
				end if   
			else
				response.write "stm_aix('p"&pNum&"i"&k&"','p"&pNum2&"i0',[0,'" & rsClass(1) & "','','',-1,-1,0,'" & rsClass(4) & "','_self','" & rsClass(4) & "','" & rsClass(6) & "','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#FFFFFF',0,'','',3,3,0,0,'#fffff7','#ff0000','#3e3e3e','#620000','']);" & vbcrlf
			end if			
		end if
		k=k+1
		rsClass.movenext
	loop
	rsClass.close
	set rsClass=nothing
	response.write "stm_ep();" & vbcrlf	
end sub

'==================================================
'过程名：ShowAnnounce
'作  用：显示本站公告信息
'        AnnounceNum  ----最多显示多少条公告
'==================================================
sub ShowAnnounce(AnnounceNum)
	dim sqlAnnounce,rsAnnounce,i
	if AnnounceNum>0 and AnnounceNum<=10 then
		sqlAnnounce="select top " & AnnounceNum
	else
		sqlAnnounce="select top 10"
	end if
	sqlAnnounce=sqlAnnounce & " * from affiche order by ID Desc"	
	Set rsAnnounce= Server.CreateObject("ADODB.Recordset")
	rsAnnounce.open sqlAnnounce,conn,1,1
	if rsAnnounce.bof and rsAnnounce.eof then 
		AnnounceCount=0
		response.write "<p>&nbsp;&nbsp;没有公告</p>" 
	else 
		AnnounceCount=rsAnnounce.recordcount		
			response.Write "本站公告："
			do while not rsAnnounce.eof   
				response.Write "&nbsp;<a href='#' onclick=""javascript:window.open('Affiche.asp?ID=" & rsAnnounce("id") &"', 'newwindow', 'height=450, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "'><font color='#FF0000'>" &rsAnnounce("title") & "</font></a>"
				rsAnnounce.movenext
				i=i+1				  
			loop       		
	end if  
	rsAnnounce.close
	set rsAnnounce=nothing
end sub

'==================================================
'过程名：ShowFriendLinks
'作  用：显示友情链接站点
'参  数：LinkType  ----链接方式，1为LOGO链接，2为文字链接
'       SiteNum   ----最多显示多少个站点
'       Cols      ----分几列显示
'       ShowType  ----显示方式。1为向上滚动，2为横向列表，3为下拉列表框
'==================================================
sub ShowFriendLinks(LinkType,SiteNum,Cols,ShowType)
	dim sqlLink,rsLink,SiteCount,i,strLink
	if LinkType<>1 and LinkType<>2 then
		LinkType=1
	else
		LinkType=Cint(LinkType)
	end if
	if SiteNum<=0 or SiteNum>100 then
		SiteNum=10
	end if
	if Cols<=0 or Cols>20 then
		Cols=10
	end if
	if ShowType=1 then'
        strLink=strLink & "<div id=rolllink style=overflow:hidden;height:100;width:100><div id=rolllink1>"    '新增加的代码
	elseif ShowType=3 then
		strLink=strLink & "<select name='FriendSite' onchange=""if(this.options[this.selectedIndex].value!=''){window.open(this.options[this.selectedIndex].value,'_blank');}""><option value=''>友情文字链接站点</option>"
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "<table width='100%' cellSpacing='5'><tr align='center' >"
	end if
	
	sqlLink="select top " & SiteNum & " * from FriendLinks where IsOK=True and LinkType=" & LinkType & " order by IsGood,id desc"
	set rsLink=server.createobject("adodb.recordset")
	rsLink.open sqlLink,conn,1,1
	if rsLink.bof and rsLink.eof then
		if ShowType=1 or ShowType=2 then
	  		for i=1 to SiteNum
				strLink=strLink & "<td>"			
				strLink=strLink & "</td>"
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' >"
				end if
			next
		end if
	else
		SiteCount=rsLink.recordcount
		for i=1 to SiteCount
			if ShowType=1 or ShowType=2 then
			  if LinkType=1 then
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='网站名称：" & rsLink("SiteName") & vbcrlf & "网站地址：" & rsLink("SiteUrl") & vbcrlf & "网站简介：" & rsLink("SiteIntro") & "'>"
				if rsLink("LogoUrl")="" or rsLink("LogoUrl")="http://" then
					strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0'>"
				else
					strLink=strLink & "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
				end if
				strLink=strLink & "</a></td>"
			  else
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='网站名称：" & rsLink("SiteName") & vbcrlf & "网站地址：" & rsLink("SiteUrl") & vbcrlf & "网站简介：" & rsLink("SiteIntro") & "'>" & rsLink("SiteName") & "</a></td>"
			  end if
			  if i mod Cols=0 and i<SiteNum then
				strLink=strLink & "</tr><tr align='center' >"
			  end if
			else
				strLink=strLink & "<option value='" & rsLink("SiteUrl") & "'>" & rsLink("SiteName") & "</option>"
			end if
			rsLink.moveNext
		next
		if SiteCount<SiteNum and (ShowType=1 or ShowType=2) then
			for i=SiteCount+1 to SiteNum
				if LinkType=1 then
					strLink=strLink & "<td width='88'></td>"
				else
					strLink=strLink & "<td width='88'></td>"
				end if
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' >"
				end if
			next
		end if
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "</tr></table>"
	end if
	if ShowType=1 then
        strLink=strLink & "</div><div id=rolllink2></div></div>"   '新增代码
	elseif ShowType=3 then
		strLink=strLink & "</select>"
	end if
	response.write strLink
	if ShowType=1 then call RollFriendLinks()    '新增代码
	rsLink.close
	set rsLink=nothing
end sub

'==================================================
'过程名：RollFriendLinks
'作  用：滚动显示友情链接站点
'参  数：无
'==================================================
sub RollFriendLinks()
%>
<script>
   var rollspeed=30
   rolllink2.innerHTML=rolllink1.innerHTML //克隆rolllink1为rolllink2
   function Marquee(){
   if(rolllink2.offsetTop-rolllink.scrollTop<=0) //当滚动至rolllink1与rolllink2交界时
   rolllink.scrollTop-=rolllink1.offsetHeight  //rolllink跳到最顶端
   else{
   rolllink.scrollTop++
   }
   }
   var MyMar=setInterval(Marquee,rollspeed) //设置定时器
   rolllink.onmouseover=function() {clearInterval(MyMar)}//鼠标移上时清除定时器达到滚动停止的目的
   rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//鼠标移开时重设定时器
</script>
<%
end sub
%>



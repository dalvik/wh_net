 
<!--#Include File="Check_Sql.asp"-->
<!--#include file="Conn.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="Function.asp"-->

<%
dim strFileName,MaxPerPage,ShowSmallClassType
dim totalPut,CurrentPage,TotalPages
dim BeginTime,EndTime
dim founderr, errmsg
dim EnBigClassName,EnSmallClassName,keyword,strField
dim rs,sql,sqlProduct,rsProduct,sqlSearch,rsSearch,sqlBigClass,rsBigClass
BeginTime=Timer
EnBigClassName=Trim(request("EnBigClassName"))
EnSmallClassName=Trim(request("EnSmallClassName"))
keyword=trim(request("keyword"))
'if keyword<>"" then 
'	keyword=replace(replace(replace(replace(keyword,"'","��"),"<","&lt;"),">","&gt;")," ","&nbsp;")
'end if
strField=trim(request("Field"))

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

sqlBigClass="select * from BigClass order by BigClassID"
Set rsBigClass= Server.CreateObject("ADODB.Recordset")
rsBigClass.open sqlBigClass,conn,1,1

'=================================================
'��������ShowSmallClass_Tree
'��  �ã�����Ŀ¼��ʽ��ʾ��Ŀ
'��  ������
'=================================================
Sub ShowSmallClass_Tree()
%>
<SCRIPT language=javascript>
function opencat(cat,img){
	if(cat.style.display=="none"){
	cat.style.display="";
	img.src="img/class2.gif";
	}	else {
	cat.style.display="none"; 
	img.src="img/class1.gif";
	}
}
</Script>
<TABLE cellSpacing=0 cellPadding=0 width="99%" border=0>
<%
dim i
set rsbig = server.CreateObject ("adodb.recordset")
		sql="select * from BigClass"
		rsbig.open sql,conn,1,1
if rsbig.eof and rsbig.bof then
	Response.Write "Columns In Building����"
else
	i=1
	do while not rsbig.eof
%>
	<TR>
		<TD language=javascript onmouseup="opencat(cat10<%=i%>000,&#13;&#10; img10<%=i%>000);" id=item$pval[catID]) style="CURSOR: hand" width=34 height=24 align=center><IMG id=img10<%=i%>000 src="img/class1.gif" width=15 height=17></TD>
		<TD width="662"><a href='EnProduct.asp?EnBigClassName=<%=rsbig("EnBigClassName")%>'><%=rsbig("EnBigClassName")%></a></TD>
	</TR>
	<TR>
		<TD id=cat10<%=i%>000 <%if rsbig("EnBigClassName")=EnBigClassName then 
		     response.write "style='DISPLAY'"   
		    else  
		     response.write "style='DISPLAY: none'"
		    end if%> colspan="2">
<%
dim rsSmall,sqls,j
set rsSmall = server.CreateObject ("adodb.recordset")
sqls="select * from SmallClass where EnBigClassName='" & rsbig("EnBigClassName") & "' order by SmallClassID"
rsSmall.open sqls,conn,1,1
if rsSmall.eof and rsSmall.bof then
	Response.Write "No SmallClass"
else
	j=1
	do while not rsSmall.eof
%>
&nbsp;<IMG height=20 src="img/class3.gif" width=26 align=absMiddle border=0><a href="EnProduct.asp?EnBigClassName=<%=rsSmall("EnBigClassName")%>&EnSmallClassName=<%=rsSmall("EnSmallClassName")%>"><%=rsSmall("EnSmallClassName")%></a><BR>
<%
	rsSmall.movenext
	j=j+1
	loop
end if
rsSmall.close
set rsSmall=nothing
%>
		</TD>
	</TR>
<%
	rsbig.movenext
	i=i+1
	loop
	rsbig.close
    set rsbig=nothing
end if
%>
</TABLE>
<%
end Sub 

'=================================================
'��������ShowVote
'��  �ã���ʾ��վ����
'��  ������
'=================================================
sub ShowVote()
	dim sqlVote,rsVote,i
	sqlVote="select top 1 * from Vote where IsSelected=True"
	Set rsVote= Server.CreateObject("ADODB.Recordset")
	rsVote.open sqlVote,conn,1,1
	if rsVote.bof and rsVote.eof then 
		response.Write "&nbsp;No research"
	else
		response.write "<form name='VoteForm' method='post' action='vote.asp' target='_blank'><td>"
		response.write rsVote("Title") & "<br><br>"
		if rsVote("VoteType")="Single" then
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='radio' name='VoteOption' value='" & i & "'>" & rsVote("Select" & i) & "<br>"
			next
		else
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='checkbox' name='VoteOption' value='" & i & "'>" & rsVote("Select" & i) & "<br>"
			next
		end if
		response.write "<br><input name='VoteType' type='hidden'value='" & rsVote("VoteType") & "'>"
		response.write "<input name='Action' type='hidden' value='Vote'>"
		response.write "<input name='ID' type='hidden' value='" & rsVote("ID") & "'>"
		response.write "<a href='javascript:VoteForm.submit();'><img src='images/voteSubmit.gif' width='52' height='18' border='0'></a>&nbsp;&nbsp;"
        response.write "<a href='Vote.asp?Action=Show' target='_blank'><img src='images/voteView.gif' width='52' height='18' border='0'></a>"
		response.write "</td></form>"
	end if
	rsVote.close
	set rsVote=nothing
end sub

'=================================================
'��������ShowClassGuide
'��  �ã���ʾ��Ŀ����λ��
'��  ������
'=================================================
sub ShowClassGuide()
	response.write  "&nbsp;<a href='EnProducts.asp'>Product</a>&nbsp;&gt;&gt;"
	if EnBigClassName="" and EnSmallClassName="" then
		response.write "&nbsp;All Products"
	else
		if EnBigClassName<>"" then
			response.write "&nbsp;<a href='EnProduct.asp?EnBigClassName=" & EnBigClassName & "'>" & EnBigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if EnSmallClassName<>"" then
				response.write "<a href='EnProduct.asp?EnBigClassName=" & EnBigClassName & "&EnSmallClassName=" & EnSmallClassName & "'>" & EnSmallClassName & "</a>"
			else
				response.write "All SmallClass"
			end if
		end if
	end if
end sub

'=================================================
'��������ShowProductTotal
'��  �ã���ʾ��������
'��  ������
'=================================================
sub ShowProductTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from Product where Passed=True "
	if EnBigClassName<>"" then
		sqlTotal=sqlTotal & " and EnBigClassName='" & EnBigClassName & "' "
		if EnSmallClassName<>"" then
			sqlTotal=sqlTotal & " and EnSmallClassName='" & EnSmallClassName & "' "
		end if	
	end if
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "Total 0 Products"
	else
		totalPut=rsTotal(0)
		response.Write "Total " & totalPut & " Products"
	end if
	rsTotal.close
	set rsTotal=nothing
end sub

'=================================================
'��������ShowProduct
'=================================================
sub ShowProduct(TitleLen)
	if TitleLen<0 or TitleLen>200 then
		TitleLen=50
	end if
    if currentpage<1 then
	   	currentpage=1
    end if
	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
	   		currentpage= totalPut \ MaxPerPage
		else
		   	currentpage= totalPut \ MaxPerPage + 1
		end if
   	end if
	if currentPage=1 then
        sqlProduct="select top " & MaxPerPage	
	else
		sqlProduct="select "
	end if

	sqlProduct=sqlProduct & " ID,Product_Id,EnBigClassName,EnSmallClassName,IncludePic,EnTitle,Price,EnSpec,EnUnit,EnMemo,DefaultPicUrl,UpdateTime,Hits from Product where Passed=True "
	
	if EnBigClassName<>"" then
		sqlProduct=sqlProduct & " and EnBigClassName='" & EnBigClassName & "' "
		if EnSmallClassName<>"" then
			sqlProduct=sqlProduct & " and EnSmallClassName='" & EnSmallClassName & "' "
		end if
	end if
	sqlProduct=sqlProduct & " order by UpdateTime desc"
	Set rsProduct= Server.CreateObject("ADODB.Recordset")
	rsProduct.open sqlProduct,conn,1,1
	if rsProduct.bof and  rsProduct.eof then
		response.Write("<br><li>No products</li>")
	else
		if currentPage=1 then
			call ProductContent(TitleLen)
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsProduct.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsProduct.bookmark
            	call ProductContent(TitleLen)
        	else
	        	currentPage=1
           		call ProductContent(TitleLen)
	    	end if
		end if
	end if
	rsProduct.close
	set rsProduct=nothing
end sub

sub ProductContent(intTitleLen)
   	dim i,strTemp
    i=0
	do while not rsProduct.eof
		strTemp=""		
		strTemp= strTemp & "<table width=100% border=0 cellspacing=3 cellpadding=0>"
                strTemp= strTemp & "<tr>"
                strTemp= strTemp & "<td width=25% rowspan=6>"
                strTemp= strTemp & "<div align=center><a href=EnProductShow.asp?ID=" & rsProduct("ID") & ">" 
				
				fileExt=lcase(getFileExtName(rsProduct("DefaultPicUrl")))
				if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then
                strTemp= strTemp & "<img style='BORDER-LEFT-COLOR: #cccccc; BORDER-BOTTOM-COLOR: #cccccc; BORDER-TOP-COLOR: #cccccc; BORDER-RIGHT-COLOR: #cccccc' src=" & rsProduct("DefaultPicUrl") & " width=105 height=80 onload='javascript:DrawImage(this);'>" 
				else
				 if fileext="swf" then
				    strTemp= strTemp & "<object  classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='105' height='84'>"
					strTemp= strTemp &"<param name=movie value='"&rsProduct("DefaultPicUrl")&"'>"
					strTemp= strTemp &"<param name=quality value=high>"
					strTemp= strTemp &"<param name='Play' value='-1'>"
					strTemp= strTemp &"<param name='Loop' value='0'>"
					strTemp= strTemp &"<param name='Menu' value='-1'>"					
					strTemp= strTemp &"<embed  src='"&rsProduct("DefaultPicUrl")&"' width='105' height='84' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed> </object>"												
			   end if
		      end if			 
				 
                strTemp= strTemp & "</a></div></td>"
                strTemp= strTemp & "<td width=18% height=12>"
                strTemp= strTemp & "Product Name��</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & "<a href=EnProductShow.asp?ID=" & rsProduct("ID") & ">" & rsProduct("EnTitle") & ""
                strTemp= strTemp & "</a></td>"						
				
				'strTemp= strTemp & "</tr><tr>" 
                'strTemp= strTemp & "<td width=12% height=12>"
                'strTemp= strTemp & "��Ʒ�ۼۣ�</td>"
                'strTemp= strTemp & "<td>" & rsProduct("Price") & "Ԫ</td>"	
				
				strTemp= strTemp & "</tr><tr>"
				strTemp= strTemp & "<td width=18% height=12>"
                strTemp= strTemp & "Product Spec��</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & rsProduct("EnSpec") & ""
                strTemp= strTemp & "</a></td>"
				
				strTemp= strTemp & "</tr><tr>"
				strTemp= strTemp & "<td width=12% height=12>"
                strTemp= strTemp & "Product Memo��</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & rsProduct("EnMemo") & ""
                strTemp= strTemp & "</a></td>"			
				
                strTemp= strTemp & "</tr><tr>"
                strTemp= strTemp & "<td height=12>"
                strTemp= strTemp & "Product Class��</td>"
                strTemp= strTemp & "<td><a href=EnProduct.asp?EnBigClassName="& rsProduct("EnBigClassName")&">"&rsProduct("EnBigClassName")&"</a> �� "
                strTemp= strTemp & "<a href=EnProduct.asp?EnBigClassName=" & rsProduct("EnBigClassName") & "&EnSmallClassName=" & rsProduct("EnSmallClassName") & ">" & rsProduct("EnSmallClassName") & ""
                strTemp= strTemp & "</a></td>"
                strTemp= strTemp & "</tr><tr>" 				

			   
                strTemp= strTemp & "<td height=12>Product Info��</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & "<a href=EnProductShow.asp?ID=" & rsProduct("ID") & " target=_blank><img src=Images/arrow_7En.gif border=0></a></td>"
                strTemp= strTemp & "</tr><tr>"
                strTemp= strTemp & "<td colspan=2>"
                strTemp= strTemp & "<table width=100% border=0 cellpadding=0 cellspacing=0>"
                strTemp= strTemp & "<tr>" 
                strTemp= strTemp & "<td width=50% height=12>"
                strTemp= strTemp & "<div align=center></div></td>"
                
				strTemp= strTemp & "<td width=50% height=12>"
                strTemp= strTemp & "<div align=center><input name='Product_Id' type='checkbox'    id='Product_Id' value="&cstr(rsProduct("Product_Id"))&"> Select"
                strTemp= strTemp & "</div></td>"
				
                strTemp= strTemp & "</tr>"
                strTemp= strTemp & "</table>"
                strTemp= strTemp & "</td>"
                strTemp= strTemp & "</tr><tr>" 
                strTemp= strTemp & "<td height=1 colspan=3 bgcolor=#CCCCCC></td>"
                strTemp= strTemp & "</tr>"
                strTemp= strTemp & "</table>"
		response.write strTemp
		rsProduct.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop
end sub 


'=================================================
'��������ShowSearchTerm
'��  �ã���ʾ����������Ϣ
'��  ������
'=================================================
sub ShowSearchTerm()
	response.write "&nbsp;Product Search&nbsp;&gt;&gt; "
	if EnBigClassName<>"" then
		response.write "<a href='EnProduct.asp?EnBigClassName=" & EnBigClassName & "'>" & EnBigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
		if EnSmallClassName<>"" then
			response.write "<a href='EnProduct.asp?EnBigClassName=" & EnBigClassName & "&EnSmallClassName=" & EnSmallClassName & "'>" & EnSmallClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
		else
			response.write "All Smallclass&nbsp;&gt;&gt;&nbsp;"
		end if
	end if
	if keyword<>"" then
	  select case strField
		case "Title"
			response.Write "Name include <font color=red>"&keyword&"</font> Products"
		case "Content"
			response.Write "Readme include <font color=red>"&keyword&"</font> Products"
		case else
			response.Write "Name include <font color=red>"&keyword&"</font> Products"
	  end select
	else
	  response.write "&nbsp;All Products"
	end if
end sub

'=================================================
'��������SearchResultTotal
'��  �ã���ʾ�����������
'��  ������
'=================================================
sub SearchResultTotal()
	dim rsTotal,sqlTotal
	sqlTotal="select count(*) from Product where Passed=True "
	if EnBigClassName<>"" then
		sqlTotal=sqlTotal & " and EnBigClassName='" & EnBigClassName & "' "
		if EnSmallClassName<>"" then
			sqlTotal=sqlTotal & " and EnSmallClassName='" & EnSmallClassName & "' "
		end if	
	end if
	if keyword<>"" then
		select case strField
			case "Title"
				sqlTotal=sqlTotal & " and EnTitle like '%" & keyword & "%' "
			case "Content"
				sqlTotal=sqlTotal & " and EnContent like '%" & keyword & "%' "
			case else
				sqlTotal=sqlTotal & " and EnTitle like '%" & keyword & "%' "
		end select
	end if
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "Total 0 Product"
	else
		totalPut=rsTotal(0)
		response.Write "Total search " & totalPut & " Products"
	end if
end sub

'=================================================
'��������ShowSearchResult
'��  �ã���ҳ��ʾ�������
'��  ������
'=================================================
sub ShowSearchResult()
    if currentpage<1 then
	   	currentpage=1
    end if
	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
	   		currentpage= totalPut \ MaxPerPage
		else
		   	currentpage= totalPut \ MaxPerPage + 1
		end if
   	end if
	if currentPage=1 then
        sqlSearch="select top " & MaxPerPage	
	else
		sqlSearch="select "
	end if

	sqlSearch=sqlSearch & " * from Product where Passed=True "
	if EnBigClassName<>"" then
		sqlSearch=sqlSearch & " and EnBigClassName='" & EnBigClassName & "' "
		if EnSmallClassName<>"" then
			sqlSearch=sqlSearch & " and EnSmallClassName='" & EnSmallClassName & "' "
		end if	
	end if
	if keyword<>"" then
		select case strField
			case "Title"
				sqlSearch=sqlSearch & " and EnTitle like '%" & keyword & "%' "
			case "Content"
				sqlSearch=sqlSearch & " and EnContent like '%" & keyword & "%' "
			case else
				sqlSearch=sqlSearch & " and EnTitle like '%" & keyword & "%' "
		end select
	end if
	sqlSearch=sqlSearch & " order by ID desc"
	Set rsSearch= Server.CreateObject("ADODB.Recordset")
	rsSearch.open sqlSearch,conn,1,1
 	if rsSearch.eof and rsSearch.bof then 
       		response.write "<p align='center'><br><br>No Search Product</p>" 
   	else 
   		if currentPage=1 then 
       		call SearchResultContent()
   		else 
       		if (currentPage-1)*MaxPerPage<totalPut then 
       			rsSearch.move  (currentPage-1)*MaxPerPage 
       			dim bookmark 
       			bookmark=rsSearch.bookmark 
       			call SearchResultContent()
      		else 
        		currentPage=1 
       			call SearchResultContent()
      		end if 
	   	end if 
   	end if 
   	rsSearch.close 
   	set rsSearch=nothing   
end sub

sub SearchResultContent()
	dim i,strTemp
    i=0
	do while not rsSearch.eof
		strTemp=""		
		strTemp= strTemp & "<table width=100% border=0 cellspacing=3 cellpadding=0>"
                strTemp= strTemp & "<tr>"
                strTemp= strTemp & "<td width=25% rowspan=6>"
                strTemp= strTemp & "<div align=center><a href=EnProductShow.asp?ID=" & rsSearch("ID") & ">" 
				
				fileExt=lcase(getFileExtName(rsSearch("DefaultPicUrl")))
				if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then
                strTemp= strTemp & "<img border=0 src=" & rsSearch("DefaultPicUrl") & " width=105 height=80>" 
				else
				 if fileext="swf" then
				    strTemp= strTemp & "<object  classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='105' height='84'>"
					strTemp= strTemp &"<param name=movie value='"&rsSearch("DefaultPicUrl")&"'>"
					strTemp= strTemp &"<param name=quality value=high>"
					strTemp= strTemp &"<param name='Play' value='-1'>"
					strTemp= strTemp &"<param name='Loop' value='0'>"
					strTemp= strTemp &"<param name='Menu' value='-1'>"
					strTemp= strTemp &"<param name='wmode' value='transparent'>"
					strTemp= strTemp &"<embed src='"&rsSearch("DefaultPicUrl")&"' width='105' height='84' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed> </object>"												
			   end if
		      end if			 
				 
                strTemp= strTemp & "</a></div></td>"
                strTemp= strTemp & "<td width=18% height=12>"
                strTemp= strTemp & "Product name:</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & "<a href=EnProductShow.asp?ID=" & rsSearch("ID") & ">" & rsSearch("EnTitle") & ""
                strTemp= strTemp & "</a></td>"
				
				'strTemp= strTemp & "</tr><tr>" 
                'strTemp= strTemp & "<td width=12% height=12>"
                'strTemp= strTemp & "��Ʒ�ۼۣ�</td>"
                'strTemp= strTemp & "<td>" & rsSearch("Price") & "Ԫ</td>"				
				
				strTemp= strTemp & "</tr><tr>"
				strTemp= strTemp & "<td width=18% height=12>"
                strTemp= strTemp & "Product Spec:</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & rsSearch("EnSpec") & ""
                strTemp= strTemp & "</a></td>"
				
				strTemp= strTemp & "</tr><tr>"
				strTemp= strTemp & "<td width=18% height=12>"
                strTemp= strTemp & "Product Memo:</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & rsSearch("EnMemo") & ""
                strTemp= strTemp & "</a></td>"				
				
                strTemp= strTemp & "</tr><tr>"
                strTemp= strTemp & "<td height=12>"
                strTemp= strTemp & "Product Class:</td>"
                strTemp= strTemp & "<td><a href=EnProduct.asp?EnBigClassName=" & rsSearch("EnBigClassName") & ">" & rsSearch("EnBigClassName") & "</a> �� "
                strTemp= strTemp & "<a href=EnProduct.asp?EnBigClassName=" & rsSearch("EnBigClassName") & "&EnSmallClassName=" & rsSearch("EnSmallClassName") & ">" & rsSearch("EnSmallClassName") & ""
                strTemp= strTemp & "</a></td>"
                strTemp= strTemp & "</tr><tr>" 

			   
                strTemp= strTemp & "<td height=12>Product Info:</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & "<a href=EnProductShow.asp?ID=" & rsSearch("ID") & "><img src=Images/arrow_7En.gif border=0></a></td>"
                strTemp= strTemp & "</tr><tr>"
                strTemp= strTemp & "<td colspan=2>"
                strTemp= strTemp & "<table width=100% border=0 cellpadding=0 cellspacing=0>"
                strTemp= strTemp & "<tr>" 
                strTemp= strTemp & "<td width=50% height=12>"
                strTemp= strTemp & "<div align=center></div></td>"
                
				strTemp= strTemp & "<td width=50% height=12>"
                strTemp= strTemp & "<div align=center><input name='Product_Id' type='checkbox'    id='Product_Id' value="&cstr(rsSearch("Product_Id"))&"> Select"
                strTemp= strTemp & "</div></td>"
				
                strTemp= strTemp & "</tr>"
                strTemp= strTemp & "</table>"
                strTemp= strTemp & "</td>"
                strTemp= strTemp & "</tr><tr>" 
                strTemp= strTemp & "<td height=1 colspan=3 bgcolor=#CCCCCC></td>"
                strTemp= strTemp & "</tr>"
                strTemp= strTemp & "</table>"
		response.write strTemp
		rsSearch.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop
end sub 


'=================================================
'��������ShowSearch
'��  �ã���ʾ������������
'��  ����ShowType ----��ʾ��ʽ��1Ϊ����2Ϊ����
'=================================================
sub ShowSearch(ShowType)
	dim count
	if ShowType<>1 and ShowType<>2 then
		ShowType=1
	end if
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
subcat[<%=count%>] = new Array("<%= trim(rs("EnSmallClassName"))%>","<%= trim(rs("EnBigClassName"))%>","<%= trim(rs("EnSmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.EnSmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.EnSmallClassName.options[document.myform.EnSmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    
</script>
<table border="0" cellpadding="2" cellspacing="0" align="center">
  <form method="Get" name="myform" action="Ensearch.asp">
    <tr>
      <td height="1">
        <%if ShowType=1 then%>
      </td>
    </tr>
  
  
    <tr>
      <td height="28"> 
        <%end if%>
        <input type="text" name="keyword"  size=12 value="keyword" maxlength="50" onFocus="this.select();"> 
        <input type="submit" name="Submit"  value="Search"> </td>
    </tr>
  </form>
</table>
<%
end sub

'=================================================
'��������ShowAllClass
'��  �ã���ʾ������Ŀ����Ŀ������
'��  ������
'=================================================
sub ShowAllClass()
	if rsBigClass.bof and rsBigClass.eof then 
		response.Write "&nbsp;No column"
	else
		dim sqlClass,rsClass,strClassName
		rsBigClass.movefirst
		do while not rsBigClass.eof
			strClassName= "��<a href='EnProduct.asp?EnBigClassName=" & rsBigClass("EnBigClassName") & "'><b>" & rsBigClass("EnBigClassName") & "</b></a>��<br><br>"
			sqlClass="select * from SmallClass where EnBigClassName='" & rsBigClass("EnBigClassName") & "' Order by SmallClassID"
			Set rsClass= Server.CreateObject("ADODB.Recordset")
			rsClass.open sqlClass,conn,1,1
			do while not rsClass.eof
				strClassName=strClassName & "&nbsp;<a href='EnProduct.asp?EnBigClassName=" & rsClass("EnBigClassName") & "&EnSmallClassName=" & rsClass("EnSmallClassName") & "'>" & rsClass("EnSmallClassName") & "</a>&nbsp;"
				rsClass.movenext
			loop
			response.write strClassName & "<br><br>"
			rsBigClass.movenext
		loop
		rsClass.close
		set rsClass=nothing
	end if
end sub


'=================================================
'��������ShowProductContent
'��  �ã���ʾ���¾�������ݣ����Է�ҳ��ʾ
'��  ������
'=================================================
sub ShowProductContent()
	dim ID,strContent,CurrentPage
	dim ContentLen,MaxPerPage,pages,i,lngBound
	dim BeginPoint,EndPoint
	ID=rs("ID")
	strContent=rs("EnContent")
	response.write strContent
end sub
'=================================================
'��������ShowUserLogin
'��  �ã���ʾ�û���¼����
'��  ������
'=================================================
sub ShowUserLogin()
	dim strLogin
	If Session("UserName")="" Then
    	strLogin= "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
		strLogin=strLogin & "<form action='EnUserLogin.asp' method='post' name='UserLogin' onSubmit='return CheckForm();'>"
        strLogin=strLogin & "<tr><td height='25' align='right'>User name��</td><td height='25'><input name='UserName' type='text' id='UserName' size='10' maxlength='20'></td></tr>"
        strLogin=strLogin & "<tr><td height='25' align='right'>Password��</td><td height='25'><input name='Password' type='password' id='Password' size='10' maxlength='20'></td></tr>"
        strLogin=strLogin & "<tr align='center'><td height='25' colspan='2'><input name='Login' type='submit' id='Login' value=' Login '> <input name='Reset' type='reset' id='Reset' value=' Reset '>"
        strLogin=strLogin & "</td></tr>"
        strLogin=strLogin & "<tr><td height='20' align='center' colspan='2'><a href='EnUserReg.asp' target='_blank'>Register</a>&nbsp;&nbsp;<a href='EnGetPassword.asp' target='_blank'>Forget password��</a></td></tr>"      
        strLogin=strLogin & "</form></table>"
		response.write strLogin
%>
<script language=javascript>
	function CheckForm()
	{
		if(document.UserLogin.UserName.value=="")
		{
			alert("Please input user name��");
			document.UserLogin.UserName.focus();
			return false;
		}
		if(document.UserLogin.Password.value == "")
		{
			alert("Please input password��");
			document.UserLogin.Password.focus();
			return false;
		}
	}
</script>
<%
	Else 
		response.write "Welcome��" & Session("UserName") & "<br><br>"
		response.write "<b>User control panel��</b><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='EnUserServer.asp'><b>User center</b></a><br><br>"
	end if
end sub
%>
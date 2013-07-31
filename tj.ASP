<!--#include file="Inc/SysProduct.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<SCRIPT language=javascript>
function LayerCh(num){
			for (i=1; i<=6; i++){
			document.all["Layer"+i].style.visibility="hidden";
				if (i==num){
				document.all["Layer"+i].style.visibility="visible";
				}
			}
		}		
		function menu2(index) {
			if (navigator.userAgent.indexOf("MSIE") != -1) {
				if (index == 1) {
					t1.style.display = "";
					t2.style.display = "none";												
				} 
				else if (index == 2) {
					t1.style.display = "none";
					t2.style.display = "";												
				}
			}
		}
		
		//????,????? ??????
		// ????? ??? ?? ??? (?????)
		var slideShowSpeed = 5500

		// ??? ?? (?)
		var crossFadeDuration = 3

		// ??? ???? ?? ???
		var Pic = new Array() 
		var imgUrl0 = 'http://image.megapass.net/relatedmaterial/Main/main/01(3).jpg';
		var imgUrl1 = 'http://image.megapass.net/relatedmaterial/Main/main/02.jpg';
		var imgUrl2 = 'http://image.megapass.net/relatedmaterial/Main/main/03.jpg';
		var imgUrl3 = 'http://image.megapass.net/relatedmaterial/Main/main/04.jpg';
		var imgUrl4 = 'http://image.megapass.net/relatedmaterial/Main/main/05.jpg';
		var imgUrl5 = '';
		var imgUrl6 = '';
		var imgUrl7 = '';
		var imgUrl8 = '';
		var imgUrl9 = '';
		if (imgUrl0 != "" )	{Pic[0] = imgUrl0;}
		if (imgUrl1 != "" )	{Pic[1] = imgUrl1;}
		if (imgUrl2 != "" )	{Pic[2] = imgUrl2;}
		if (imgUrl3 != "" )	{Pic[3] = imgUrl3;}
		if (imgUrl4 != "" )	{Pic[4] = imgUrl4;}
		if (imgUrl5 != "" )	{Pic[5] = imgUrl5;}
		if (imgUrl6 != "" )	{Pic[6] = imgUrl6;}
		if (imgUrl7 != "" )	{Pic[7] = imgUrl7;}
		if (imgUrl8 != "" )	{Pic[8] = imgUrl8;}
		if (imgUrl9 != "" )	{Pic[9] = imgUrl9;}
		//Pic[3] = '../img4.jpg'
		//Pic[4] = '../img5.jpg'

		var t
		var j = 0
		var p = Pic.length

		var preLoad = new Array()
		for (i = 0; i < p; i++){
		preLoad[i] = new Image()
		preLoad[i].src = Pic[i]
		}

		function runSlideShow(){
			if (document.all){
				document.images.SlideShow.style.filter="blendTrans(duration=2)"
				document.images.SlideShow.style.filter="blendTrans(duration=crossFadeDuration)"
				document.images.SlideShow.filters.blendTrans.Apply()      
			}

			document.images.SlideShow.src = preLoad[j].src

			if (document.all){
				document.images.SlideShow.filters.blendTrans.Play()
			}

			j = j + 1
			if (j > (p-1)) j=0
			t = setTimeout('runSlideShow()', slideShowSpeed)
		}
		
		function img_href() {
			if(document.all){
				if(document.images.SlideShow.src == imgUrl0) {location.href='http://tvclick.megapass.net/ktweb/newSpecial/free_detail_contents.aspx?section=81&contentsid=28643&idx=1'}
				if(document.images.SlideShow.src == imgUrl1) {location.href='http://tvclick.megapass.net/ktweb/newSpecial/free_detail_contents.aspx?section=81&contentsid=27324&idx=2'}
				if(document.images.SlideShow.src == imgUrl2) {location.href='http://tvclick.megapass.net/ktweb/newSpecial/free_detail_contents.aspx?section=81&contentsid=32386&idx=5'}
				if(document.images.SlideShow.src == imgUrl3) {location.href='http://tvclick.megapass.net/ktweb/newSpecial/free_detail_contents.aspx?section=81&contentsid=32146&idx=6'}
				if(document.images.SlideShow.src == imgUrl4) {location.href='http://tvclick.megapass.net/ktweb/newSpecial/free_detail_contents.aspx?section=83&contentsid=42562&idx=1'}
				if(document.images.SlideShow.src == imgUrl5) {location.href=''}
				if(document.images.SlideShow.src == imgUrl6) {location.href=''}
				if(document.images.SlideShow.src == imgUrl7) {location.href=''}
				if(document.images.SlideShow.src == imgUrl8) {location.href=''}
				if(document.images.SlideShow.src == imgUrl9) {location.href=''}
			}
			//location.href = var1;
		}
		</script> <DIV id=Layer1 
            style="Z-INDEX: 6; LEFT: 0px; VISIBILITY: visible; POSITION: absolute; TOP: 7px">
            <TABLE cellSpacing=0 cellPadding=0 width=540 border=0>
              <!--DWLayoutTable-->
              <TBODY>
              <TR>
                <TD width=80 height="29" align=right vAlign=top>
                  <TABLE cellSpacing=0 cellPadding=0 width=75 border=0>
                    <TBODY>
                    <TR>
                      <TD style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop2_01.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(2) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_02.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(3) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_03.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(4) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_04.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(5) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_05.gif" 
                        border=0></TD>
						<TD 
                        style="CURSOR: hand" onmouseover=javascript:LayerCh(6)><IMG 
                        src="images/newtop/newtop1_06.gif" 
                        border=0></TD></TR></TBODY></TABLE></TD>
             
                </TR>
              <%Set rs = Server.CreateObject("ADODB.Recordset")
sql="select Content from Aboutus where Title='企业简介'"
rs.open sql,conn,1,3%>
<table width="571" border="0" cellspacing="0" cellpadding="0" class="border2" height="221">
               <tr>
               <TD width="10">
               
</TD> <TD height="190" width="570">
                <%=rs("content")%>
</TD><TD width="10">
               
</TD></tr></table>
             <%rs.close
rs=nothing%>   
              </TBODY></TABLE></TD>
                <TD></TD>
                </TR></TBODY></TABLE>
</DIV>

            <DIV id=Layer2 
           style="Z-INDEX: 5; LEFT: 0px; VISIBILITY: hidden; POSITION: absolute; TOP: 7px">
             <TABLE cellSpacing=0 cellPadding=0 width=540 border=0>
              <!--DWLayoutTable-->
              <TBODY>
              <TR>
               
			    <TD width=80 height="29" align=right vAlign=top>
                  <TABLE cellSpacing=0 cellPadding=0 width=75 border=0>
                    <TBODY>
                    <TR>
                      <TD onmouseover=javascript:LayerCh(1) style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_01.gif" 
                        border=0></TD>
                      <TD style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop2_02.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(3) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_03.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(4) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_04.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(5) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_05.gif" 
                        border=0></TD>
						<TD 
                        style="CURSOR: hand" onmouseover=javascript:LayerCh(6)><IMG 
                        src="images/newtop/newtop1_06.gif" 
                        border=0></TD></TR></TBODY></TABLE></TD>
              <td width="1"></td>
                </TR>
               
               <tr>
                <TD width=580 height="190" valign="top"><SPAN id=mainWeekbest2>
               <table width="100%" border="0" cellspacing="0" cellpadding="0" class="border2" height="221">
              
                <% rs.open "select top 20 * from news where BigClassName='企业新闻' order by AddDate desc ",conn,1,1
					  i=-1
					 do while not rs.eof
				
					  %>

               
                  
                    <td onMouseOver="this.style.backgroundColor='#EEF7FF';return true;" onMouseOut="this.style.backgroundColor='#ffffff';" width="25" height="18"><img src="images/icon_new.gif" width="18" height="15"></td>
                  <td width="279" height="18"><a href="shownews.asp?id=<%=rs("id") %>" target="_blank"><font color="#000000"><%= mid(rs("title"),1,32)%></font></a> </td>
           
                <%if i mod 2 <>0 then
		end if
		if i mod 2 =0 then
		Response.Write "</TR><TR>"
		end if
		rs.MoveNext
	i=i+1
	Loop

rs.close%>
                 <tr>
                  <td>&nbsp;</td>
                  <td>　</td>
                 
                </tr>
            </table>
                <TR>
                      <TD height=5></TD></TR>
               </TBODY></TABLE></SPAN></TD>
                <TD></TD>
                </TR></TBODY></TABLE>
</DIV>
            <DIV id=Layer3 
            style="Z-INDEX: 4; LEFT: 0px; VISIBILITY: hidden; POSITION: absolute; TOP: 7px">
             <TABLE cellSpacing=0 cellPadding=0 width=540 border=0 height="221">
              <!--DWLayoutTable-->
              <TBODY>
              <TR>
                <TD width=80 height="29" align=right vAlign=top>
                  <TABLE cellSpacing=0 cellPadding=0 width=75 border=0>
                    <TBODY>
                    <TR>
                      <TD onmouseover=javascript:LayerCh(1) style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_01.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(2) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_02.gif" 
                        border=0></TD>
                      <TD style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop2_03.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(4) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_04.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(5) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_05.gif" 
                        border=0></TD>
						<TD 
                        style="CURSOR: hand" onmouseover=javascript:LayerCh(6)><IMG 
                        src="images/newtop/newtop1_06.gif" 
                        border=0></TD></TR></TBODY></TABLE></TD>
             <td width="1"></td>
                </TR>
               
              <tr>
                <TD width=580 height="221" valign="top"><SPAN id=mainWeekbest2>
               <table width="100%" border="0" cellspacing="0" cellpadding="0" class="border2" height="221">
               
                <% rs.open "select top 20 * from news where BigClassName='业内资讯' order by AddDate desc ",conn,1,1
					  i=-1
					 do while not rs.eof
				
					  %>

               
                
                    <td onMouseOver="this.style.backgroundColor='#EEF7FF';return true;" onMouseOut="this.style.backgroundColor='#ffffff';" width="25" height="18"><img src="images/icon_new.gif" width="18" height="15"></td>
                  <td width="279" height="18"><a href="shownews.asp?id=<%=rs("id") %>" target="_blank"><font color="#000000"><%= mid(rs("title"),1,32)%></font></a> </td>
           
                <%if i mod 2 <>0 then
		end if
		if i mod 2 =0 then
		Response.Write "</TR><TR>"
		end if
		rs.MoveNext
	i=i+1
	Loop

rs.close%>
                <tr>
                  <td>&nbsp;</td>
                  <td>　</td>
                 
                </tr>
            </table>
                <TR>
                      <TD height=5></TD></TR>
               </TBODY></TABLE></SPAN></TD>
                <TD></TD>
                </TR></TBODY></TABLE>
</DIV>
            <DIV id=Layer4 
            style="Z-INDEX: 3; LEFT: 0px; VISIBILITY: hidden; POSITION: absolute; TOP: 7px">
              <TABLE cellSpacing=0 cellPadding=0 width=540 border=0>
              <!--DWLayoutTable-->
              <TBODY>
              <TR>
                <TD width=80 height="29" align=right vAlign=top>
                  <TABLE cellSpacing=0 cellPadding=0 width=75 border=0>
                    <TBODY>
                    <TR>
                      <TD onmouseover=javascript:LayerCh(1) style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_01.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(2) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_02.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(3) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_03.gif" 
                        border=0></TD>
                      <TD style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop2_04.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(5) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_05.gif" 
                        border=0></TD>
						<TD 
                        style="CURSOR: hand" onmouseover=javascript:LayerCh(6)><IMG 
                        src="images/newtop/newtop1_06.gif" 
                        border=0></TD></TR></TBODY></TABLE></TD>
    
                </TR>
               
                 <tr>
                <TD width=580 height="190" valign="top"><SPAN id=mainWeekbest2>
               <table width="100%" border="0" cellspacing="0" cellpadding="0" class="border2" height="221">
           
                <% rs.open "select top 20 * from news where BigClassName='技术文档' order by AddDate desc ",conn,1,1
					  i=-1
					 do while not rs.eof
				
					  %>

               
                    <td onMouseOver="this.style.backgroundColor='#EEF7FF';return true;" onMouseOut="this.style.backgroundColor='#ffffff';" width="25" height="18"><img src="images/icon_new.gif" width="18" height="15"></td>
                  <td width="279" height="18"><a href="shownews.asp?id=<%=rs("id") %>" target="_blank"><font color="#000000"><%= mid(rs("title"),1,32)%></font></a> </td>
           
                <%if i mod 2 <>0 then
		end if
		if i mod 2 =0 then
		Response.Write "</TR><TR>"
		end if
		rs.MoveNext
	i=i+1
	Loop

rs.close%>
                <tr>
                  <td>&nbsp;</td>
                  <td>　</td>
                 
                </tr>
            </table>
                <TR>
                      <TD height=5></TD></TR>
                </TBODY></TABLE></SPAN></TD>
                <TD></TD>
                </TR></TBODY></TABLE>
</DIV>
            <DIV id=Layer5 
            style="Z-INDEX: 2; LEFT: 0px; VISIBILITY: hidden; POSITION: absolute; TOP: 7px">
              <TABLE cellSpacing=0 cellPadding=0 width=540 border=0>
              <!--DWLayoutTable-->
              <TBODY>
              <TR>
                <TD width=80 height="29" align=right vAlign=top>
                  <TABLE cellSpacing=0 cellPadding=0 width=75 border=0>
                    <TBODY>
                    <TR>
                      <TD onmouseover=javascript:LayerCh(1) style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_01.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(2) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_02.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(3) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_03.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(4) style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_04.gif" 
                        border=0></TD>
                      <TD 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop2_05.gif" 
                        border=0></TD>
						<TD 
                        style="CURSOR: hand" onmouseover=javascript:LayerCh(6)><IMG 
                        src="images/newtop/newtop1_06.gif" 
                        border=0></TD></TR></TBODY></TABLE></TD>
            
                </TR>
               
               <tr>
                <TD width=580 height="190" valign="top"><table width=570 border=0 cellSpacing=1 cellpadding=1 bgColor=#bbbbbb><%set rs=server.createobject("adodb.recordset")
   sql="select top 8 * from download order by id desc"
   	rs.open sql,conn,1,1
if rs.bof and rs.eof then
		response.Write("<br><li>没有任何下载</li>")
	else
	id=rs("id")
   title=rs("title")
   DownloadUrl=rs("DownloadUrl")
do while not rs.eof%>

 
<tr Class=tdbg><td width=60% height=18 bgColor=#FFFFFF>
<a href="DownloadShow.asp?ID=<%=ID%>" target=_blank>&nbsp;<%=rs("Title")%></a>
</td>
<td width=40% height=22 bgColor=#FFFFFF>&nbsp
<a href="DownloadShow.asp?id=<%=ID%>" target=_blank><FONT COLOR=#EC5E00>详细</FONT></a>|<a href="DownloadShow.asp?ID=<%=ID%>" target=_blank><FONT COLOR=#EC5E00>下载</FONT></a></td>
</tr>

            
     <%rs.movenext
		
	loop     
	rs.close
	rs=nothing
	end if %> </table> <SPAN id=mainWeekbest5>
               <td height="100" colspan="2"> 
                       
                <tr> 
                            <td height="8"></td>
                            <td></td>
                </tr>
              </table></td>
                    <TR>
                      <TD height=5></TD></TR>
                </TBODY></TABLE></SPAN></TD>
                <TD></TD>
                </TR></TBODY></TABLE>
</DIV>
				 <DIV id=Layer6 
            style="Z-INDEX: 1; LEFT: 0px; VISIBILITY: hidden; POSITION: absolute; TOP: 7px">
              <TABLE cellSpacing=0 cellPadding=0 width=540 border=0 height="221">
              <!--DWLayoutTable-->
              <TBODY>
              <TR>
                <TD width=80 height="29" align=right vAlign=top>
                  <TABLE cellSpacing=0 cellPadding=0 width=75 border=0>
                    <TBODY>
                    <TR>
                      <TD onmouseover=javascript:LayerCh(1) style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_01.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(2) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_02.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(3) 
                        style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_03.gif" 
                        border=0></TD>
                      <TD onmouseover=javascript:LayerCh(4) style="CURSOR: hand"><IMG 
                        src="images/newtop/newtop1_04.gif" 
                        border=0></TD>
                      <TD 
                        style="CURSOR: hand" onmouseover=javascript:LayerCh(5)><IMG 
                        src="images/newtop/newtop1_05.gif" 
                        border=0></TD>
						 <TD 
                        style="CURSOR: hand" ><IMG 
                        src="images/newtop/newtop2_06.gif" 
                        border=0></TD></TR></TBODY></TABLE></TD>
             
                </TR>
              
                <TD width=580 height="221" valign="top"><table  width="570" border="0" cellpadding="1" cellspacing="1" bgcolor="#bbbbbb" >
                      <tr Class=tdbg > 
                        <td  width="34%" height="29"> 
                          <div align="center">产品名称</font></div></td>
                            <td  width="18%"> 
                              <div align="center">规格</div></td>
                            <td  width="9%"> 
                              <div align="center">单位</div></td>
                            <td  width="39%"> 
                              <div align="center">备注</div></td>
                      </tr>    </table><table width=570 border=0 cellSpacing=1 cellpadding=1 bgColor=#bbbbbb><%i=1 
set rs_Product=server.createobject("adodb.recordset")
sqltext="select top 8 * from Product where Passed=True And Elite=True order by UpdateTime desc"
rs_Product.open sqltext,conn,1,1				 	
%>
                      <%
do while not rs_Product.eof
%>
                      <% if i mod 2 =0 then%>
                      <tr Class=tdbg> 
                        <% else %>
                      <tr bgcolor="#ffffff"> 
                          <% end if %>
                          <td height="22" width="34%"><a href="ProductShow.asp?ID=<%=rs_Product("ID")%>" target="_blank"><%=rs_Product("Title")%></a></td>
                            <td width="18%"> <div align="center"><%=rs_Product("Spec")%></div></td>
                            <td width="9%"> <div align="center"><%=rs_Product("Unit")%></div></td>
                            <td width="39%"> <div align="left"><%=rs_Product("Memo")%></div></td>
                      </tr>
                      <% 
i=i+1
rs_Product.movenext 
loop 
rs_Product.close  
%>  </table> <SPAN id=mainWeekbest5>
               <td height="100" colspan="2"> 
                       
                <tr> 
                            <td height="8"></td>
                            <td></td>
                </tr>
              </table></td>
                    <TR>
                      <TD height=5></TD></TR>
                </TBODY></TABLE></SPAN></TD>
                <TD></TD>
                </TR></TBODY></TABLE>
</DIV> 
		
           
			        
				  </body>


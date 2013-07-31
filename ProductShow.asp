<!--#include file="Inc/SysProduct.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="Head.asp" -->

<script language="JavaScript"> 
<!-- 
var flag=false; 
function DrawImage(ImgD){ 
 var image=new Image(); 
 image.src=ImgD.src; 
 if(image.width>0 && image.height>0){ 
  flag=true; 
  if(image.width/image.height>= 500/500){ 
   if(image.width>500){
    ImgD.width=500; 
    ImgD.height=(image.height*500)/image.width; 
   }else{ 
    ImgD.width=image.width;
    ImgD.height=image.height; 
   } 
   ImgD.alt="点击查看详细信息..."; 
  } 
  else{ 
   if(image.height>500){
    ImgD.height=500; 
    ImgD.width=(image.width*500)/image.height; 
   }else{ 
    ImgD.width=image.width;
    ImgD.height=image.height; 
   } 
   ImgD.alt="点击查看详细信息..."; 
  } 
 }
}
//--> 
</script> 

<%
ShowSmallClassType=ShowSmallClassType_Article
dim ID
ID=trim(request("ID"))
if ID="" then
	response.Redirect("Product.asp")
end if

sql="select * from Product where ID=" & ID & ""
Set rsshow= Server.CreateObject("ADODB.Recordset")
rsshow.open sql,conn,1,3
if rsshow.bof and rsshow.eof then
	response.write"<SCRIPT language=JavaScript>alert('找不到此产品！');"
  response.write"javascript:history.go(-1)</SCRIPT>"
else	
	rsshow("Hits")=rsshow("Hits")+1
	rsshow.update
%>
<body>

<DIV id=index_container>
  <div id="index_left">
    <div id="left">
      <div id="left_title">产品系列</div>
      <div id="left_content">
        <DIV id=elem-FrontProductCategory_slideTree-001 name="商品分类"><LINK 
href="imgbyy/css2.css" type=text/css 
rel=stylesheet>
          <div class="FrontProductCategory_slideTree-default" 
id="FrontProductCategory_slideTree-001">
            <div class="left_bottom">
              <div class="right_bottom">
                <div class="left_top">
                  <div class="right_top">
                    <table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td height="3"></td>
                      </tr>
                    </table>
                    <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td><% call ShowSmallClass_Tree() %>
                        </td>
                      </tr>
                    </table>
                    <table width="100" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td height="10"></td>
                      </tr>
                    </table>
                  </div>
                </div>
              </div>
            </div>
          </div>
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
  <DIV id=right_title>产品展示</DIV>
<DIV class=text><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="452" 
                  height=25><%
response.write "&nbsp;<a href='Product.asp?BigClassName=" & rsshow("BigClassName") & "'>" & rsshow("BigClassName") & "</a>&nbsp;&gt;&gt;&nbsp;"
if rsshow("SmallClassName") & ""<>"" then
	response.write "<a href='Product.asp?BigClassName=" & rsshow("BigClassName")&"&SmallClassName=" & rsshow("SmallClassName") & "'>" & rsshow("SmallClassName") & "</a>&nbsp;&gt;&gt;&nbsp;"
end if
response.write rsshow("Title")
%>                    </td>
                    <td width="98" valign="middle"><a href="Payment.asp?Product_Id=<%=rsshow("Product_Id")%>",target="_blank"><img border=0 src=Img/addtocart.gif
width=100 height=26></a></td>
                  </tr>
                 
                  <tr>
                    <td height="237" colspan="3"><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="21">&nbsp;</td>
                        </tr>
                        <tr>
                          <%fileExt=lcase(getFileExtName(rsshow("DefaultPicUrl")))%>
                          <td width="52%" height="200" align="center"><font color="#FF6600">&nbsp;</font><font color="#FF6600">&nbsp;</font>
                              <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                              <a href="<%=rsshow("DefaultPicUrl")%>" target="_blank"><img src=<%=rsshow("DefaultPicUrl")%>  border="0" ></a>
                             
                              <%
		end if%>                          </td>
                        </tr>
                        
                        <tr>
                          <td height="9">&nbsp;</td>
                        </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td class="title_right" height="37" colspan="3">&nbsp;&nbsp; <span style="font-weight: bold">产 品 说 明 </span></td>
                  </tr>
                  <tr>
                    <td height="1" colspan="3"><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="50"><%=rsshow("content")%>                          </td>
                        </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td 
                  height=1 colspan="3" bgcolor="#CCCCCC"></td>
                  </tr>
                  <tr>
                    <td 
                  height=40 colspan="2" align="center">点击数：<%=rsshow("Hits")%>&nbsp; 录入时间：<%= FormatDateTime(rsshow("UpdateTime"),2) %>&nbsp;【<a href='javascript:window.print()'>打印此页</a>】&nbsp;【<a href="javascript:self.close()">关闭</a>】<a href="Payment.asp?Product_Id=<%=rsshow("Product_Id")%>",target="_blank"></a></td>
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
<%
end if
rsshow.close
set rsshow=nothing
call CloseConn()
%>
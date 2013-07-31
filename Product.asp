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
  if(image.width/image.height>= 105/80){ 
   if(image.width>105){
    ImgD.width=105; 
    ImgD.height=(image.height*105)/image.width; 
   }else{ 
    ImgD.width=image.width;
    ImgD.height=image.height; 
   } 
   ImgD.alt="点击查看详细信息..."; 
  } 
  else{ 
   if(image.height>80){
    ImgD.height=80; 
    ImgD.width=(image.width*80)/image.height; 
   }else{ 
    ImgD.width=image.width;
    ImgD.height=image.height; 
   } 
   ImgD.alt="点击查看详细信息..."; 
  } 
 }
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
  }
//--> 
</script> 
<%
MaxPerPage=MaxPerPage_Default
strFileName="Product.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName 
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
<DIV class=text>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="226" 
                  height="30"><% call ShowClassGuide() %>
      </td>
      <td width="227">&nbsp;</td>
      <td width="102"><% call ShowProductTotal() %>
      </td>
    </tr>
    <tr>
      <td 
                  height="42" colspan="3"><form action="Payment.asp" method="post" name="Inquire" target="_blank" id="Inquire" >
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr valign="middle" bgcolor="#F7F7F7">
              <td width="53%" height="30"><div align="center"> </div></td>
              <td width="22%"><div align="right">
                  <input name="chkAll" type="checkbox" id="chkAll" onClick="CheckAll(this.form)" value="checkbox" />
                选取所有</div></td>
              <td width="25%" align="center"><input name="image" type="image" src="Images/inquire_now.gif" style="border: 0px "width="100" height="26" border="0" />
              </td>
            </tr>
            <tr>
              <td colspan="3" height="5"></td>
            </tr>
            <tr>
              <td colspan="3" ><% call ShowProduct(32) %>
              </td>
            </tr>
          </table>
      </form></td>
    </tr>
    <tr>
      <td 
                  height="1" colspan="3"><%
		  if totalput>0 then		    
		  	call showpage(strFileName,totalput,MaxPerPage,false,true,"个产品")
		  end if
		  %>
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

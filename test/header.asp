<!--#INCLUDE FILE="common.inc" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
</head>
<BODY>
<table width="760" border="0" align="center">
  <tr>
    <td width="738"> 
      <table width="760" border="0" align="center" height="60">
        <tr> 
          <td width="151"> 
            <div align="left"><a href="<%=PageBaseHref%>" target="_top"><img height=60 width=150 src="<%=LogoImgLocation%>" border="0"></a></div>
          </td>
          <td width="523"> 
            <div align="right"><img height=60 width=518 src="<%=BannerImgLocation%>" align="left"></div>
          </td>
          <td width="72">
            <p align="center"><font size="2"><a href="#" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('<%=PageBaseHref%>')">设为首页</a></font></p>
            <p align="center"><font size="2"><a href="#" onClick="javascript:window.external.AddFavorite('<%=PageBaseHref%>','<%=PageTitle%>')">添加收藏</a></font></p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="738">
      <table width="760" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td background="gif/gray.gif" height="25" width="10">&nbsp;</td>
          <td background="gif/gray.gif" height="25" width="214">
          	<font size='2'>欢迎你，<% = getUserName(session("user_id")) %>！</font>
          </td>
          <td background="gif/graydc.gif" height="25" width="137">&nbsp;</td>
          <td background="gif/grayline.gif" height="25" width="399" align="right" valign="top"> 
            <table border=0 cellpadding=0 cellspacing=0 height=17 width="100%">
              <tbody> 
              <tr> 
                <td align=middle width=7 height="19"> 
                  <div align="center"><font size="2">|</font></div>
                </td>
                <td align=middle width=50 height="19"> 
                  <div align="center"><font size="2"><a href="home.asp" target="fbody">新闻</a></font></div>
                </td>
                <td align=middle width=7 height="19"> 
                  <div align="center"><font size="2">|</font></div>
                </td>
                <td align=middle width=50 height="19"> 
                  <div align="center"><font size="2"><a href="address.asp" target="fbody">通讯录</a></font></div>
                </td>
                <td align=middle width=7 height="19"> 
                  <div align="center"><font size="2">|</font></div>
                </td>
                <td align=middle width=50 height="19"> 
                  <div align="center"><font size="2"><a href="album.asp" target="fbody">影集</a></font></div>
                </td>
                <td align=middle width=7 height="19"> 
                  <div align="center"><font size="2">|</font></div>
                </td>
                <td align=middle  width="50" height="19" > 
                  <div align="center"><font size="2"><a href="forum.asp" target="fbody">论坛</a></font></div>
                </td>
                <td align=middle  width="7" height="19" > 
                  <div align="center"><font size="2">|</font></div>
                </td>
                <% if getUserLevel(session("user_id"))=3 then %>
                <td align=middle  width="50" height="19" > 
                  <div align="center"><font size="2"><a href="manage.asp" target="fbody">管理</a></font></div>
                </td>
                <td align=middle  width="7" height="19" > 
                  <div align="center"><font size="2">|</font></div>
                </td>
                <% end if %>
                <td align=middle  width="50" height="19" > 
                  <div align="center"><font size="2"><a href="about.asp" target="fbody">关于</a></font></div>
                </td>
                <td align=middle  width="7" height="19" > 
                  <div align="center"><font size="2">|</font></div>
                </td>
                <td align=middle  width="25" height="19" >&nbsp;</td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>


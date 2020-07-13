<!--#INCLUDE FILE="config.inc" -->

<html>
<head>
<title><% =PageTitle %></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
<bgsound src="heavens.mid" loop="-1">
</head>
<body>
<table width="605" border="0" cellpadding="10" align="center" height="398">
  <tr> 
    <td height="12" width="660"></td>
  </tr>
  <tr> 
    <td height="52" width="660"> 
      <div align="center"><font face="<% =SpecificFontFace %>" size="5" color="#0000FF"><b><% =PageTitle %></b></font></div>
    </td>
  </tr>
  <tr> 
    <td height="31" width="660"> 
      <p> 
      <div align="left"><font face="<% =SpecificFontFace %>" size="4" color="#FF6633">&nbsp;&nbsp;&nbsp;&nbsp;早已习惯了年少时轻率的分别与咫尺天涯的相隔。音书不闻，了无消息。<br>
        &nbsp;&nbsp;&nbsp;&nbsp;拿着手上临别时的赠言，愈看觉得愈不可信起来。时间使白纸黑字也陈旧起来，老式的联系方式也让感情与信念僵化起来。<br>
        &nbsp;&nbsp;&nbsp;&nbsp;这过去便在脑海里退居至无人的角落中了。终于在这里我找到了那一切。<br>
        &nbsp;&nbsp;&nbsp;&nbsp;想留住曾经的相聚，或涉水越山找寻远去的踪迹，终于在那个多雾的清晨慢慢模糊了曾如此熟悉和清晰的背影。 
        <br>
        &nbsp;&nbsp;&nbsp;&nbsp;想采撷固有的思念，或赠予那渐行渐远的决意的过客，终于将艰涩的想念化成每一次习惯的查询。</font> 
      </div>
    </td>
  </tr>
  <tr> 
    <td height="166" width="660"> 
      <form method="post" action="login.asp" target="_top">
        <table width="100%" border="0" align="center" height="87">
          <tr> 
            <td width="40%">
              <div align="center"><font face="<% =DefaultFontFace %>" >姓名:</font>
                <input type="text" name="name">
              </div>
            </td>
            <td width="40%">
              <div align="center"><font face="<% =DefaultFontFace %>" >密码: </font>
                <input type="password" name="password">
              </div>
            </td>
            <td width="20%"> 
              <div align="center">
                           <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="50" align="center">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:document.forms[0].submit()"><font size=2>登录</font></a></div>
                		</td>
              		      </tr>
            		   </table>              
              </div>
            </td>
          </tr>
          <tr valign="top"> 
            <td height="44" colspan="3"> 
              <hr>
              <font face="<% =DefaultFontFace %>" size="-1">提示：请在姓名栏输入你的中文姓名，初始密码为“<%=SiteName%>”。<br>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;如果你有什么问题，请发邮件到<a href="mailto:<%=Webmaster%>"><%=Webmaster%></a>。 
              <br>
              </font> </td>
          </tr>
          <tr valign="bottom"> 
            <td height="37" colspan="3"> 
              <div align="center"><b><font face="<% =DefaultFontFace %>" color="#0000FF">你是第 
                <!--#INCLUDE FILE="counter.asp" -->
                位访问者</font></b></div>
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
</body>
</html>

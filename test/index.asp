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
      <div align="left"><font face="<% =SpecificFontFace %>" size="4" color="#FF6633">&nbsp;&nbsp;&nbsp;&nbsp;����ϰ��������ʱ���ʵķֱ���������ĵ���������鲻�ţ�������Ϣ��<br>
        &nbsp;&nbsp;&nbsp;&nbsp;���������ٱ�ʱ�����ԣ�����������������������ʱ��ʹ��ֽ����Ҳ�¾���������ʽ����ϵ��ʽҲ�ø����������������<br>
        &nbsp;&nbsp;&nbsp;&nbsp;���ȥ�����Ժ����˾������˵Ľ������ˡ��������������ҵ�����һ�С�<br>
        &nbsp;&nbsp;&nbsp;&nbsp;����ס��������ۣ�����ˮԽɽ��ѰԶȥ���ټ����������Ǹ�������峿����ģ�����������Ϥ�������ı�Ӱ�� 
        <br>
        &nbsp;&nbsp;&nbsp;&nbsp;���ߢ���е�˼��������ǽ��н�Զ�ľ���Ĺ��ͣ����ڽ���ɬ�������ÿһ��ϰ�ߵĲ�ѯ��</font> 
      </div>
    </td>
  </tr>
  <tr> 
    <td height="166" width="660"> 
      <form method="post" action="login.asp" target="_top">
        <table width="100%" border="0" align="center" height="87">
          <tr> 
            <td width="40%">
              <div align="center"><font face="<% =DefaultFontFace %>" >����:</font>
                <input type="text" name="name">
              </div>
            </td>
            <td width="40%">
              <div align="center"><font face="<% =DefaultFontFace %>" >����: </font>
                <input type="password" name="password">
              </div>
            </td>
            <td width="20%"> 
              <div align="center">
                           <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="50" align="center">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:document.forms[0].submit()"><font size=2>��¼</font></a></div>
                		</td>
              		      </tr>
            		   </table>              
              </div>
            </td>
          </tr>
          <tr valign="top"> 
            <td height="44" colspan="3"> 
              <hr>
              <font face="<% =DefaultFontFace %>" size="-1">��ʾ�������������������������������ʼ����Ϊ��<%=SiteName%>����<br>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������ʲô���⣬�뷢�ʼ���<a href="mailto:<%=Webmaster%>"><%=Webmaster%></a>�� 
              <br>
              </font> </td>
          </tr>
          <tr valign="bottom"> 
            <td height="37" colspan="3"> 
              <div align="center"><b><font face="<% =DefaultFontFace %>" color="#0000FF">���ǵ� 
                <!--#INCLUDE FILE="counter.asp" -->
                λ������</font></b></div>
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
</body>
</html>

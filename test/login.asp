<!--#include file="config.inc"-->

<%
'On Error Resume Next

username=Request.form("name")
password=Request.form("password")

set conn=server.createobject("adodb.connection")
conn.open ConnString

strSql ="SELECT Member_id ,M_Last_Visited from Members where M_Name = '" & userName & "' and M_Password = '" & password &"'"
'Response.Write StrSql
set rs = conn.Execute (StrSql)

'��¼�û���¼���
if rs.BOF or rs.EOF then
	logtype = "�Ƿ�"
else
	logtype = "�Ϸ�"
end if
str = "insert into Login (L_Name, L_Type, L_IP) Values ('"
str = str & username & "', '"
str = str & logtype & "', '"
str = str & Request.ServerVariables("REMOTE_ADDR") & "')"
'Response.write str
conn.Execute (str)

if rs.BOF or rs.EOF then
%>
	<html>
	<head>
		<TITLE>��¼����</title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<link rel=stylesheet type=text/css href='style.css'>
	</head>
	<body>
	<center><br><br>
		<h2><font face="<% =SpecificFontFace %>" color="<% =DefaultFontColor %>">��¼����</font></h2>
		<br><font face="<% =DefaultFontFace %>" size="3">
		<p>�ʺš�<font color="<% =SpecificFontColor %>"><% =username %></font>�������ڣ������������������
		<br><br>������������롣</p>
		</font>
		<br>
		<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        <tr> 
          <td> 
            <div align="center" class="p9"><a href="<% =PageBaseHref %>"><font size=2>����</font></a></div>
          </td>
        </tr>
      </table>
    </center>
	</body>
	</html>
<%	
Else
	'## save user infomation
	session("user_id") = rs("Member_id")
	session("last_visited") = rs("M_Last_Visited")
	
	'refresh visit time
	strSql = "update Members set M_Last_Visited = #" & Now() & "#, M_Visited_Times = M_Visited_Times+1 where M_Name = '" & userName & "'"
	conn.Execute (StrSql)
	
	'redirect to main.asp
	Response.redirect("main.asp")
End if
rs.close
set rs=nothing
	
'## �ر����ݿ�����
conn.close
set conn=nothing
%>

<!--#INCLUDE FILE="common.inc" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv='Refresh' content='60; url=userlist.asp'>
<link rel=stylesheet type=text/css href='style.css'>
</head>
<body>
<div align="right">
  <table width="90" border="0" cellspacing="0" cellpadding="0" style="border: 1px solid rgb(0,0,0)">
    <tr> 
      <td height="25" bgcolor="#9999FF" align="center"><font size="2">在线用户</font></td>
    </tr>    
<%
set conn=server.createobject("adodb.connection")
conn.open ConnString

defDate = dateadd("s", -60, now)

strSql ="SELECT * from Members where M_Last_Visited >#" & defDate & "# order by M_Name"    
'Response.Write StrSql
set rs = conn.Execute (StrSql)

i=0
do until rs.Eof 
	if i mod 2= 0 then 
		CColor = "#CCFFFF"
	else
		CColor = "#CCCCFF"
	End if
	
	Response.Write "<tr>"
	Response.Write "<td bgcolor='" & CColor & "' height='25'><div align='center'><font size='2' color='#FF00FF' face='楷体_GB2312'><a href=address.asp?id=" & rs("Member_id") & " target='fbody'>" & rs("M_Name") & "</a></font></div></td>"
	Response.Write "</tr>"
		
	i=i+1
	rs.MoveNext
loop

'## 关闭数据库连接
rs.close
set rs=nothing

conn.close
set conn=nothing
%>
  </table>
</div>
</body>
</html>
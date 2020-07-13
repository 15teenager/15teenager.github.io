<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout.asp"
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
</head>
<body>
<%
set conn=server.createobject("adodb.connection")
conn.open ConnString

if session("news_id") <> "" then
	newsid = session("news_id")	
else
	Response.Redirect "home.asp"
end if
	
strSql ="SELECT News.N_Posted_By from News where News_ID = " & newsid 
'Response.Write StrSql
set rs = conn.Execute (StrSql)

If rs.Eof or rs.Bof then  
	GO_Result "新闻不存在!", false, "javascript:history.back()"
Elseif rs("N_Posted_By")= Session("user_id") or getUserLevel(session("user_id"))=3 then 
	strSQL = "Delete * From News where News_ID = " & newsid
	conn.Execute strSQL
	GO_Result "成功删除新闻!", true, "home.asp"
Else	
	GO_Result "你没有权限删除这条新闻!", false, "javascript:history.back()"
End if

rs.close
set rs=nothing

'## 关闭数据库连接
conn.close
set conn=nothing
%>	
</body>
</html>
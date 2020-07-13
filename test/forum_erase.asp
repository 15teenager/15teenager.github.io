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

if session("topic_id") <> "" then
	topicid = session("topic_id")	
else
	Response.Redirect "forum.asp"
end if
	
strSql ="SELECT Topics.T_Originator, Topics.T_ParentID from Topics where Topic_ID = " & topicid 
'Response.Write StrSql
set rs = conn.Execute (StrSql)

If rs.Eof or rs.Bof then  
	GO_Result "帖子不存在!", false, "javascript:history.back()"
Elseif rs("T_Originator")= Session("user_id") or getUserLevel(session("user_id"))=3 then 
	strSql = "SELECT * FROM Topics where Topics.T_ParentID = " &  topicid
	set rs1 = conn.Execute (strSql)
	
	If rs1.Eof or rs1.Bof then 
		strSQL = "Delete * From Topics where Topic_ID = " & topicid
		conn.Execute strSQL
		
		'# Update Last Post
		rootid=GetTopicID(topicid)
		strSql = "update Topics set T_Replies = T_Replies-1 where Topic_ID = " & rootid
		conn.Execute (StrSql)
		
		GO_Result "成功删除帖子!", true, "forum_info.asp?topic_id=" & rs("T_ParentID")
	Else
		GO_Result "已有跟帖，不能删除这个帖子!", false, "javascript:history.back()"
	End if
	
	rs1.close
	set rs1=nothing
Else	
	GO_Result "你没有权限删除这个帖子!", false, "javascript:history.back()"
End if

rs.close
set rs=nothing

'## 关闭数据库连接
conn.close
set conn=nothing
%>
</body>
</html>

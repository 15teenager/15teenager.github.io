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

photoid=request.Form("photo_id")
	
strSql ="SELECT * from Album INNER JOIN Photoes ON Album.Album_ID = Photoes.P_Album_ID  where Photoes.Photo_ID = " & photoid 
'Response.Write StrSql
set rs = conn.Execute (StrSql)

If rs.Eof or rs.Bof then  
	GO_Result "照片不存在!", false, "javascript:history.back()"
Else
	if chkPermission(rs("P_Album_ID"))=true then
		dir=Server.URLEncode(rs("A_Directory"))
		fn=Server.URLEncode(rs("P_Filename"))
		
		strSQL = "Delete * From Photoes where Photo_ID = " & photoid
		conn.Execute strSQL
		
		if Err.description <> "" then 
			GO_Result "数据库错误： " & Err.description, false, "javascript:history.back()"
		Else
			Response.redirect ("del.cgi?path=" & dir & "&file=" & fn)
			'Response.write ("del.cgi?path=" & dir & "&file=" & fn)
			'Go_Result  "照片删除成功！", true, "album_info.asp"
		End If
	else
		GO_Result "你没有权限删除这张照片!", false, "javascript:history.back()"
	end if
End if

rs.close
set rs=nothing

'## 关闭数据库连接
conn.close
set conn=nothing
%>
</body>
</html>
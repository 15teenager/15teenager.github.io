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

if session("album_id") <> "" then
	albumid=session("album_id")	
else
	Response.Redirect "album_info.asp"
end if

strSql ="SELECT * from Album INNER JOIN Photoes ON Album.Album_ID = Photoes.P_Album_ID  where Album.Album_ID = " & albumid 
'Response.Write StrSql
set rs = conn.Execute (StrSql)

If rs.Eof or rs.Bof then  
	if chkPermission(albumid)=true and getAlbumOwner(albumid)=0 then
		strSql ="SELECT * from Album where Album.Album_ID = " & albumid 
		'Response.Write StrSql
		set rs2 = conn.Execute (StrSql)
		
		if Not rs2.Eof then
			dir=Server.URLEncode(rs2("A_Directory"))
			
			strSQL = "Delete * From Album where Album_ID = " & albumid
			conn.Execute strSQL
		
			if Err.description <> "" then 
				GO_Result "���ݿ���� " & Err.description, false, "javascript:history.back()"
			Else
				Response.redirect ("del.cgi?path=" & dir)
				'Response.write ("del.cgi?path=" & dir)
				'Go_Result  "Ӱ��ɾ���ɹ���", true, "javascript:window.close()"
			End If
		end if
		
		rs2.close
		set rs2=nothing
	else
		GO_Result "��û��Ȩ��ɾ�����Ӱ��!", false, "javascript:history.back()"
	end if
Else
	GO_Result "Ӱ���л�����Ƭ������ɾ��!", false, "javascript:history.back()"
End if

rs.close
set rs=nothing

'## �ر����ݿ�����
conn.close
set conn=nothing
%>
</body>
</html>
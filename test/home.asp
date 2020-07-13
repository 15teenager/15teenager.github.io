<%
if Session("user_id") = "" then
	Response.Redirect "timeout.asp"
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<frameset cols="120,*" border=0 frameborder=0 framespacing=0 framemargin=0>
    <frame name=fleft framespacing=0 frameborder=0 marginwidth=0 marginheight=0 noresize src="userlist.asp">
    <frame name=fmain framespacing=0 frameborder=0 marginwidth=0 marginheight=0 src="news.asp">
</frameset>
</html>
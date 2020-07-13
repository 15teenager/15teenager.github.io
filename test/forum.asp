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
<table width=760 align=center border=0 cellspacing=2 cellpadding=0>
  <tr><td align=right>
     <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
       <tr> 
         <td> 
           <div align="center" class="p9"><a href=forum_post.asp><font size=2>发帖子</font></a></div>
         </td>
       </tr>
     </table>	
  </td></tr>
</table>
<br>
<%
set conn=server.createobject("adodb.connection")
conn.open ConnString

mypage=request("pageno")
If  mypage="" then
	if session("forum_pageno") <> "" then
		mypage = session("forum_pageno")	
	else
   		mypage=1
   	end if
end if
session("forum_pageno")=mypage

'##     打开数据库连接
strSql ="SELECT * FROM Topics where T_ParentID = 0 order by T_Last_Posted DESC"

set rs=server.createobject("adodb.recordset")
rs.cachesize=Pagesize
rs.open strSql,conn,3
%>
<table border="0" width="760" align="center" cellspacing="2" cellpadding="0">
  <tr>
    <td align="center" bgcolor="<% =HeadCellColor %>">&nbsp;</td>
    <td align="center" bgcolor="<% =HeadCellColor %>"><strong><font face="<% =DefaultFontFace %>" size="2" color="<% =HeadFontColor %>">主题</font></strong></td>
    <td align="center" bgcolor="<% =HeadCellColor %>"><strong><font face="<% =DefaultFontFace %>" size="2" color="<% =HeadFontColor %>">作者</font></strong></td>
    <td align="center" bgcolor="<% =HeadCellColor %>"><strong><font face="<% =DefaultFontFace %>" size="2" color="<% =HeadFontColor %>">回复</font></strong></td>
    <td align="center" bgcolor="<% =HeadCellColor %>"><strong><font face="<% =DefaultFontFace %>" size="2" color="<% =HeadFontColor %>">最新帖子</font></strong></td>
  </tr>
<% 
'显示所有Topic
If rs.Eof or rs.Bof then  
	' No items found in DB
	Response.Write "<tr><td>&nbsp;</td><td collspan=4>没有帖子！</td></tr>"
Else
	rs.movefirst
	
	rs.pagesize=Pagesize
	maxpages=cint(rs.pagecount)
	if cint(mypage) > maxpages then 
		mypage=maxpages
		session("forum_pageno")=mypage
	end if
	rs.absolutepage=mypage
	
	i=0
	rec = 1
	do until rs.Eof or rec > Pagesize '## Display Forum
		if i = 0 then 
			CColor = AltForumCellColor
		else
			CColor = ForumCellColor
		End if
	  
	  	Response.Write "<tr>"
	  	Response.Write "<td width=20 bgcolor='" & CColor & "' align='center'>" & isNew(rs("T_Last_Posted"),4) & "</td>" & vbcrlf
	  	Response.Write "<td bgcolor='" & CColor & "'><a href='forum_info.asp?topic_id=" & rs("Topic_ID") & "'><img src='gif/old_T.gif' align='middle' border=0><font face='" & DefaultFontFace & "' size='2'>" & rs("T_Subject") & "</font></a></td>"
	  	Response.Write "<td width=60 bgcolor='" & CColor & "' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'><a href=address.asp?id=" & rs("T_Originator") & ">" &  getUserName(rs("T_Originator")) & "</a></font></td>"
	  	Response.Write "<td width=30 bgcolor='" & CColor & "' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>" & rs("T_Replies") & "</font></td>"
	  	Response.Write "<td width=140 bgcolor='" & CColor & "' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='1'>" & rs("T_Last_Posted") & "</font></td>"
	  	Response.Write "</tr>"
	  	rs.MoveNext
	  	
	    	i = i + 1
	    	if i = 2 then 
	    		i = 0 
	    	end if
	    	rec = rec + 1
	loop
End If

rs.close
set rs=nothing
%>
</table>  
<%if maxpages > 1 then  %>
<hr width="98%" size="2" align="center">
<table width="93%" align="center">
<tr><td>
<div align=left><font face="<% =DefaultFontFace %>" size="2">
<%
	pad="&nbsp;&nbsp;"
	scriptname=request.servervariables("script_name")
	Response.Write "页数: &nbsp;&nbsp; " 
	for counter=1 to maxpages
   		If counter>10 then
      			pad="&nbsp;"
	   	end if
   
   		if counter <> cint(mypage) then   
			ref="<a href='" & scriptname 
			ref=ref & "?pageno=" & counter
			ref=ref & "'>" & counter & "</a>"
			response.write ref & pad
   		Else
   			Response.Write "<font color='" & DefaultFontColor & "'>" & counter & "</font>" & pad
   		End if
   
   
   		if counter mod Linesize = 0 then
      			response.write "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
   		end if
	next
%>
</font></div>
</td></tr></table>
<%
End if

'## 关闭数据库连接
conn.close
set conn=nothing
%>	       
</body>
</html>
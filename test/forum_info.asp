<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout.asp"
end if
%>

<%
'this function is displaying the child messages
sub displayMessages(layer,id)
	dim sql_order,rs_order,str,MessageSpacing,spaceing
	
	'find child with his parent id
	sql_order = "SELECT * FROM Topics where Topics.T_ParentID = " &  id 

	set rs_order = conn.Execute (sql_order)
	
	do until rs_order.eof
		str=""
		spaceing=""
		for MessageSpacing=1 to layer
			spaceing=spaceing & "&nbsp;&nbsp;&nbsp;"
		next
		
		str= "<tr><td><font face='DefaultFontFace' size='2'>" & spaceing 
		if rs_order("Topic_ID")=cint(session("topic_id")) then
			str=str & "<img src='gif/file.gif' border=0>" & rs_order("T_subject")
		else
			str=str & "<a href='forum_info.asp?topic_id=" & rs_order("Topic_ID") & "'><img src='gif/file.gif' border=0>" & rs_order("T_subject") & "</a>" 
		end if
		str=str & " .... " & rs_order("T_date") & " .... <a href=address.asp?id=" & rs_order("T_Originator") & ">" & getUserName(rs_order("T_Originator")) & "</a>"
		str=str & isNew(rs_order("T_date"),5) & "</font></td></tr>"
		
		'printing details
		Response.Write str 
		'calling again to find more childs - if not rs_order=nothing
	    	call displayMessages(layer+1,rs_order("topic_id"))
		rs_order.movenext
	loop
	
	'closing object
	rs_order.close
	set rs_order=nothing
End Sub	
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
<script language="JavaScript">
<!--
function doConfirm()
{
  if (window.confirm("确实要删除这个帖子吗？"))
    document.forms[0].submit();
}
// -->
</script>
</head>
<body>
<%
set conn=server.createobject("adodb.connection")
conn.open ConnString

topicid=request.QueryString("topic_id")
If  topicid="" then
	if session("topic_id") <> "" then
		topicid = session("topic_id")	
	else
   		Response.Redirect "forum.asp"
   	end if
   
end if
session("topic_id")=topicid

'##     打开数据库连接
strSql = "SELECT * FROM Topics where topics.topic_id = " &  topicid 

set rs = conn.Execute (StrSql)
If rs.Eof or rs.Bof then  ' No categories found in DB
	Response.Redirect "forum.asp"
Else
%>
<form method="POST" action="forum_erase.asp">
<table width=760 align=center border=0 cellspacing=2 cellpadding=0>
	<tr>
	<td align=left>
		<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="left">
              	  <tr> 
                    <td> 
                      <div align="center" class="p9"><a href="forum_post.asp?method=reply"><font size=2>回复帖子</font></a></div>
                    </td>
              	  </tr>
            	</table>
	</td>
	<td width=80 align=right>
		<table width=80> 
                    <tr>
                      <% if rs("T_Originator")= Session("user_id") or getUserLevel(session("user_id"))=3 then %>
                      <td width="33%">
      			<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        		  <tr> 
          		    <td> 
            		      <div align="center" class="p9"><a href=forum_post.asp?method=edit><font size=2>修改帖子</font></a></div>
          		    </td>
        		  </tr>
      			</table>
    		      </td>
    		      <td width="33%"> 
      			<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        		  <tr> 
          		    <td> 
            		      <div align="center" class="p9"><a href="javascript:doConfirm()"><font size=2>删除帖子</font></a></div>
          		    </td>
        		  </tr>
      			</table>
    		      </td>
    		      <%end if%>
    		      <td width="33%"> 
      			<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        		  <tr> 
          		    <td> 
            		      <div align="center" class="p9"><a href="forum.asp"><font size=2>返回</font></a></div>
          		    </td>
        		  </tr>
      			</table>
    		      </td>
    		    </tr>
    		  </table>
	</td></tr>
</table>
</form>
<table border="0" width="760" align=center cellspacing="2" cellpadding="0">
  <tr>
    <td align="center" bgcolor="<% =HeadCellColor %>"  width="60"><strong><font face="<% =DefaultFontFace %>" size="2" color="<% =HeadFontColor %>">作者</font></strong></td>
    <td align="center" bgcolor="<% =HeadCellColor %>"><strong><font face="<% =DefaultFontFace %>" size="2" color="<% =HeadFontColor %>">主题</font></strong></td>
  </tr>
<%
	Response.Write "<tr>"
	Response.Write "<td width='60' bgcolor='" & ForumCellColor & "' valign='top' align='center'><font color='" & ForumFontColor & "' face='" & DefaultFontFace & "' size='2'><a href=address.asp?id=" & rs("T_Originator") & ">" & getUserName(rs("T_Originator")) & "</a></font></td>"
	Response.Write "<td bgcolor='" & ForumCellColor & "'  valign='top' ><font color='" & ForumFontColor & "' face='" & DefaultFontFace & "' size='2'>标题： " &  "<font face='" & SpecificFontFace & "'>"  & rs("T_subject") & "</font>"
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;发表于 - " & formatDate(rs("T_Date")) & "</font>"
	Response.Write "<hr noshade size=1><font color='" & ForumFontColor & "' face='" & DefaultFontFace & "' size='2'>" &  formatStr(rs("T_Message")) & "</font></td>"
	Response.Write "</tr></TD></TR>"
End If

rs.close
set rs=nothing
%> 
</table>     
<hr width="98%" size="2" align="center">
<div align=center>
<table width=760 border=0 cellspacing=0 cellpadding=0>
<%
'find parent id
strSql = "SELECT * FROM Topics where topics.topic_id = " &  GetTopicID(topicid)

set rs = conn.Execute (StrSql)

'printing details
str= "<tr><td><font face='DefaultFontFace' size='2'>"
if rs("Topic_ID")=cint(session("topic_id")) then
	str=str & "<img src='gif/file.gif' border=0>" & rs("T_subject")
else
	str=str & "<a href='forum_info.asp?topic_id=" & rs("Topic_ID") &"'><img src='gif/file.gif' border=0>" & rs("T_subject") & "</a>"
end if
str=str & " .... " & rs("T_date") & " .... <a href=address.asp?id=" & rs("T_Originator") & ">" & getUserName(rs("T_Originator")) & "</a>"
str=str & isNew(rs("T_date"),5) & "</font></td></tr>"
Response.Write str 
	
'calling procedure to show childs
call displayMessages(1,rs("topic_id"))

rs.close
set rs=nothing
%>
</table>  
</div>
<%
'## 关闭数据库连接
conn.close
set conn=nothing
%>
</body>
</html>


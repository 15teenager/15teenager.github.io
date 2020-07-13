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
<script language="JavaScript">
<!--
function doConfirm()
{
  if (window.confirm("确实要删除这条新闻吗？"))
    document.forms[0].submit();
}

function WinOpen2(url) {
	var AtWnd = window.open(url,"AlbumWindow","toolbar=yes,directories=no,menubar=yes,location=no,status=yes,scrollbars=yes,resizable=yes,copyhistory=no");
	if ((document.AttachForm.window != null) && (!AtWnd.opener)) {
		AtWnd.opener = document.AttachForm.window;
		AtWnd.focus();
	}
}
// -->
</script>
</head>
<body>
<%
set conn=server.createobject("adodb.connection")
conn.open ConnString

newsid=request.QueryString("news_id")
If  newsid="" then
	if session("news_id") <> "" then
		newsid = session("news_id")	
	else
   		Response.Redirect "home.asp"
   	end if
   
end if
session("news_id")=newsid

'##     打开数据库连接
strSql ="SELECT * FROM News where News.News_ID = " & newsid
set rs = conn.Execute (StrSql)

If rs.Eof or rs.Bof then  
	Response.Write "<div align=center>记录不存在！</div>"
Else
%>
<table width="760" border="0" cellspacing="2" cellpadding="3" align="center" style="border: 1px solid rgb(0,0,0)" bgcolor="<% =TableBgColor %>">
  <tr>
    <td>
      <table width=95% align="center">    
	<tr><th><br><font color=#05006C size=4><% =rs("N_Title") %></font></th></tr>
	<tr><td height=><hr size=1 bgcolor=#d9d9d9></td></tr>
	<tr><td height=20 align=center><font size=2><% =formatDate(rs("N_Posted_Date")) %>&nbsp;&nbsp;&nbsp;&nbsp;<a href=address.asp?id=<%=rs("N_Posted_By")%>><% =getUserName(rs("N_Posted_By")) %></a></font></td></tr>
	<tr><td><br>
	<p><% =formatStr(rs("N_Content")) %></p>
	<p>&nbsp;</p>
	</td></tr>
	<tr><td><br><br>
	<form method="POST"  name="AttachForm" action="news_erase.asp">
	<table width=100% border=0 cellspacing=0 cellpadding=0>
	<tr>
	<td width=15% align=left valign=top>
		<% if rs("N_Album_ID")<>0 then %>
			<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
              		  <tr> 
                	    <td> 
                  	      <div align="center" class="p9"><a href='#' onClick='WinOpen2("album_info.asp?album_id=<%=rs("N_Album_ID")%>")'><font size=2>相关影集</font></a></div>
                	    </td>
              		  </tr>
            		</table>
		<% end if %>
	</td>
	<td width=45%>&nbsp;</td>
	<td width=40% align=right>
		<table width=80> 
                    <tr>
                      <% if rs("N_Posted_By")= session("user_id") or getUserLevel(session("user_id"))=3 then %>
                      <td width="33%">
      			<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        		  <tr> 
          		    <td> 
            		      <div align="center" class="p9"><a href=news_post.asp?method=edit><font size=2>修改新闻</font></a></div>
          		    </td>
        		  </tr>
      			</table>
    		      </td>
    		      <td width="33%"> 
     			<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        		  <tr> 
          		    <td> 
            		      <div align="center" class="p9"><a href="javascript:doConfirm()"><font size=2>删除新闻</font></a></div>
          		    </td>
        		  </tr>
      			</table>
    		      </td>
    		      <%end if%>
    		      <td width="33%"> 
      			<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        		  <tr> 
          		    <td> 
            		      <div align="center" class="p9"><a href="home.asp"><font size=2>返回</font></a></div>
          		    </td>
        		  </tr>
      			</table>
    		      </td>
    		    </tr>
    		</table>
	</td></tr>
	</table>
	</form>
	</td></tr>
      </table>
  </td></tr>
</table>
<%
End if
rs.close
set rs=nothing

'## 关闭数据库连接
conn.close
set conn=nothing
%>      
</body>
</html>	

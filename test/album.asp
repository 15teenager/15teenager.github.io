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
%>
<Form name="AttachForm">
<table width="760" border="0" align="center" height="320" cellspacing="3" cellpadding="0" style="border: 1px solid rgb(0,0,0)" height="320" bgcolor="<% =TableBgColor %>">
        <tr valign="top" align="center"> 
          <td><br>
      <table width="90%" border="0">
        <tr> 
          <td height="30" width="17%" bgcolor="#FFCCCC"> 
            <div align="center"><strong><font face="<% =DefaultFontFace %>" size="2" color="#008080">集体影集</font></strong></div>
          </td>
          <td width="83%">
          <% if getUserLevel(session("user_id"))>1 then %>
      	  <table width="100%"><tr><td align=right>
      	    <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
              <tr> 
                <td> 
                  <div align="center" class="p9"><a href=album_post.asp><font size=2>创建新影集</font></a></div>
                </td>
              </tr>
            </table>
      	  </td></tr></table>
      	  <%End if%>
          </td>
        </tr>
        <tr> 
          <td colspan="2">
	    <table width="100%" border="0" cellspacing="2" cellpadding="2" align="center" style="border: 1px solid rgb(0,0,0)">

<%		
			strSql ="SELECT Album.Album_ID, Album.A_Name, Album.A_Posted_Date "
			strSql = strSql & "FROM Album where Album.A_Personal = false "
			strSql = strSql & "order by Album.A_Posted_Date Desc"
			'Response.write strSql
			set rs = conn.Execute (strSql)

			'显示所有集体影集
			If rs.Eof or rs.Bof then  
				' No items found in DB
				Response.Write "<tr><td width='30'>&nbsp;</td><td collspan=4>没有集体影集！</td></tr>"
			Else
				rs.movefirst
			
				do until rs.Eof
					Response.Write "<tr>"	
	  				Response.Write "<td width='30' align='center'>" & isNew(rs("A_Posted_Date"),3) & "</td>" & vbcrlf
	  				Response.Write "<td width='240' align='left'><a href='#' onClick='WinOpen2(""album_info.asp?album_id=" & rs("Album_ID") & """)'><img src='gif/green-ball.gif' border=0><font size='2'>" & rs("A_Name") & "</font></a></td>"
	  				Response.Write "<td width='23'>&nbsp;</td>"
		  			
		  			rs.MoveNext
	  				if rs.Eof then
	  					Response.Write "<td width='30'>&nbsp;</td><td width='240'>&nbsp;</td>"
	  				else
		  				Response.Write "<td width='30' align='center'>" & isNew(rs("A_Posted_Date"),3) & "</td>" & vbcrlf
		  				Response.Write "<td width='240' align='left'><a href='#' onClick='WinOpen2(""album_info.asp?album_id=" & rs("Album_ID") & """)'><img src='gif/green-ball.gif' border=0><font size='2'>" & rs("A_Name") & "</font></a></td>"
	  					rs.MoveNext
	  				end if
	  				Response.Write "<tr>"	
	  			loop
		  	End if
	  	
		  	rs.close
			set rs=nothing
%>
 
            </table>
          </td>
        </tr>
      </table>
      <br><br>
      <table width="90%" border="0">
        <tr> 
          <td height="30" width="17%" bgcolor="#FFCCCC"> 
            <div align="center"><strong><font face="<% =DefaultFontFace %>" size="2" color="#008080">个人影集</font></strong></div>
          </td>
          <td width="83%">&nbsp;</td>
        </tr>
        <tr> 
          <td colspan="2">
	    <table width="100%" border="0" cellspacing="2" cellpadding="2" align="center" style="border: 1px solid rgb(0,0,0)">

<%		
			strSql ="SELECT Album.Album_ID, Album.A_Name, Album.A_Posted_Date "
			strSql = strSql & "FROM Album where Album.A_Personal = true "
			strSql = strSql & "order by Album.A_Name"
			'Response.write strSql
			set rs = conn.Execute (strSql)

			'显示所有个人影集
			If rs.Eof or rs.Bof then  
				' No items found in DB
				Response.Write "<tr><td width='30'>&nbsp;</td><td collspan=4>没有个人影集！</td></tr>"
			Else
				rs.movefirst
			
				do until rs.Eof
					Response.Write "<tr>"	
	  				Response.Write "<td width='30' align='center'>" & isNew(rs("A_Posted_Date"),3) & "</td>" & vbcrlf
	  				Response.Write "<td width='240' align='left'><a href='#' onClick='WinOpen2(""album_info.asp?album_id=" & rs("Album_ID") & """)'><img src='gif/green-ball.gif' border=0><font size='2'>" & rs("A_Name") & "</font></a></td>"
		  			Response.Write "<td width='23'>&nbsp;</td>"
	  			
		  			rs.MoveNext
	  				if rs.Eof then
	  					Response.Write "<td width='30'>&nbsp;</td><td width='240'>&nbsp;</td>"
	  				else
		  				Response.Write "<td width='30' align='center'>" & isNew(rs("A_Posted_Date"),3) & "</td>" & vbcrlf
		  				Response.Write "<td width='240' align='left'><a href='#' onClick='WinOpen2(""album_info.asp?album_id=" & rs("Album_ID") & """)'><img src='gif/green-ball.gif' border=0><font size='2'>" & rs("A_Name") & "</font></a></td>"
	  					rs.MoveNext
	  				end if
	  				Response.Write "<tr>"	
	  			loop
		  	End if
	
		  	rs.close
			set rs=nothing	  	  	
%>
            </table>
          </td>
        </tr>
      </table><br> 	
    </td>
  </tr>
</table>
</form>
<%
'## 关闭数据库连接
conn.close
set conn=nothing
%>
</body>
</html>

	






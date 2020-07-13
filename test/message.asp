<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout2.asp"
end if
%>

<%
set conn=server.createobject("adodb.connection")
conn.open ConnString

'refresh last visit time
if Session("user_id") <> "" then
	str = "update Members set M_Last_Visited = #" & Now() & "# where Member_id = " & session("user_id")
	conn.Execute (Str)
end if

If Request.ServerVariables("Request_Method")="GET" then
%>
<html>
<title>发送讯息</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
<%
	mode=request.QueryString("mode")
	Select Case mode
		case 1 '发送讯息%>
<body>			
<form method="post" action="message.asp">
  <table width="450" border="0" align="center" cellspacing="2" cellpadding="0">
    <tr> 
      <td colspan="3" height="17">&nbsp;</td>
    </tr>
    <tr> 
      <th colspan="3" height="47" valign="top"><font face="<% =SpecificFontFace %>">发送讯息</font></th>
    </tr>
    <tr> 
      <td colspan="3" height="27">接收者: &nbsp;&nbsp;<input type="text" value="<%= getUserName(request.QueryString("id")) %>" size="5" disabled>
      </td>
    </tr>
    <tr> 
      <td colspan="3" height="27">讯息内容: <input type="text" name="msg" size="50" maxlength="60">
      </td>
    </tr>
    <tr> 
      <td width="45%" align="right" valign="bottom" height="48"> 
                          <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:document.forms[0].submit()"><font size=2>确 认</font></a></div>
                		</td>
              		      </tr>
            		   </table>      
      </td>
      <td width="15%" height="28">&nbsp;</td>
      <td width="45%" align="left" valign="bottom" height="48"> 
                          <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="left">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:window.close()"><font size=2>关 闭</font></a></div>
                		</td>
              		      </tr>
            		   </table>
      </td>
    </tr>
  </table>
  <input name="id" type="hidden" value="<% =request.QueryString("id") %>"> 
  <input name="m" type="hidden" value="<% =request.QueryString("mode") %>">
</form>
</body>
<%		case 2 '回复讯息%>
<body>
<form method="post" action="message.asp">
  <table width="760" border="0" align="center" cellspacing="2" cellpadding="0">
    <tr> 
      <td><font size=2>回复讯息给<font face="<% =SpecificFontFace %>" color="<% =SpecificFontColor %>"><% =getUserName(request.QueryString("id")) %></font></font>: 
        <input type="text" name="msg" size="60" maxlength="60">
      </td>
      <td width="56"> 
      <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        <tr> 
          <td> 
            <div align="center" class="p9"><a href="javascript:document.forms[0].submit()"><font size=2>发送</font></a></div>
          </td>
        </tr>
      </table>      
      </td>
    </tr>
  </table>
  <input name="id" type="hidden" value="<% =request.QueryString("id") %>"> 
  <input name="m" type="hidden" value="<% =request.QueryString("mode") %>">
</form>
</body>
<%		case Else '显示讯息
			strSql = "select * from Messages where M_ReceiveID = " & Session("user_id") & " order by M_Send_Date Desc"
			'Response.write strsql
			set rs = conn.Execute (StrSql)
			
			if rs.bof and rs.eof then
%>
<meta http-equiv='Refresh' content='30; url=message.asp'>
<%			else				
				sid=rs("M_SendID")
				msg=rs("M_Message")
				sdate=formatDate(rs("M_Send_Date"))
				
				strSQL = "Delete * From Messages where Messages_ID = " & rs("Messages_ID")
				conn.Execute strSQL
%>
<body>
<table width="760" border="0" align="center" cellspacing="2" cellpadding="0">
  <tr> 
    <td><font size=2><font face="<% =SpecificFontFace %>" color='<% =SpecificFontColor %>'><% =getUserName(sid) %></font>（<% =sdate %>）： <% =msg %></font></td>
    <td width="56"> 
      <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="50" align="center">
        <tr> 
          <td> 
            <div align="center" class="p9"><a href="message.asp?mode=2&id=<% =sid %>"><font size=2>回复</font></a></div>
          </td>
        </tr>
      </table>
    </td>
    <td width="56"> 
      <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="50" align="center">
        <tr> 
          <td> 
            <div align="center" class="p9"><a href="message.asp"><font size=2>忽略</font></a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
<%			end if
			
			rs.close
			set rs=nothing
	End Select%>
</html>
<%
Else
	id=request.Form("id")
	msg=trim(request.Form("msg"))
	
	if id<>"" and msg<>"" then
		strSql = "insert into Messages (M_SendID, M_ReceiveID, M_Message) Values ("
		strSql = StrSql & Session("user_id") & ", "
		strSql = StrSql & id & ", '"
		strSql = StrSql & msg & "')"
		'Response.write strsql
		conn.Execute (StrSql)
	end if
	
	if Request.Form("m")=1 then
		Response.write "<script language='javascript'>window.close()</script>"
	else
		Response.Redirect "message.asp"
	end if
End if

'## 关闭数据库连接
conn.close
set conn=nothing
%>




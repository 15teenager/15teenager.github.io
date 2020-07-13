<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout2.asp"
end if
%>

<%
Function EmailDropDown(emails, name)
	Response.Write "<Select Name='" & name & "'>" & vbcrlf
	
	if emails="" then
		Response.Write "<Option>无</option>"  & vbcrlf
	else
		mail = split(emails, ";")
		for i = 0 to ubound(mail)
			Response.Write "<option>" & mail(i) & "</option>" & vbcrlf
		next
	end if
	
	Response.Write "</select>" & vbcrlf
End Function	
%>

<html>
<head>
<title>发送邮件</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
</head>
<body>
<%set conn=server.createobject("adodb.connection")
conn.open ConnString
%>
<table width="450" height="310" align="center" border="0" cellspacing="2" cellpadding="0"> 
  <tr valign="top"> 
    <td>
<%If Request.ServerVariables("Request_Method")="GET" then%>    
      <form method="post" action="mail.asp">
        <table width="100%" border="0" align="center">
          <tr> 
      	    <th colspan="2"><font face="<% =SpecificFontFace %>">发送邮件</font></th>
    	  </tr>
    	  <tr> 
      	    <td colspan="2">&nbsp;</td>
    	  </tr>
          <tr> 
            <td height="25" width="17%">发送：</td>
            <td height="25" width="83%"><input type="text" size="5" value="<% =getUserName(session("user_id")) %>" disabled>
               <% =EmailDropDown(getEmail(session("user_id")), "smail") %>
            </td>
          </tr>
          <tr> 
            <td height="25" width="17%">接收：</td>
            <td height="25" width="83%"><input type="text" size="5" value="<% =getUserName(request.QueryString("id")) %>" disabled> 
              <% =EmailDropDown(getEmail(request.QueryString("id")), "rmail") %>
            </td>
          </tr>
          <tr> 
            <td height="25" width="17%">主题：</td>
            <td height="25" width=83%"> 
              <input type="text" name="subject" size="52">
            </td>
          </tr>
          <tr> 
            <td width=17%" valign="top">正文：</td>
            <td width=83%" valign="top"> 
              <textarea name="content" cols="52" rows="8"></textarea>
            </td>
          </tr>
          <tr> 
            <td height="60" colspan="2" valign="top"> 
              <table width="100%" border="0">
                <tr> 
                  <td width="45%" align="right" valign="middle" height="58"> 
                    <table border="1" cellspacing="0" cellpadding="2" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
                      <tr> 
                        <td> 
                          <div align="center" class="p9"><A href="javascript:document.forms[0].submit()"><font size=2>确 认</font></a></div>
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td width="15%" height="58">&nbsp;</td>
                  <td width="45%" align="left" valign="middle" height="58"> 
                    <table border="1" cellspacing="0" cellpadding="2" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="left">
                      <tr> 
                        <td> 
                          <div align="center" class="p9"><A href="javascript:window.close()"><font size=2>关 闭</font></a></div>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
  	       </table>
            </td>
          </tr>
        </table>
      </form>
<%Else
   	Set JMail = Server.CreateObject("JMail.SMTPMail")
   	JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
  	'JMail.ServerAddress = "mail.abc.org"
  	If request.Form("smail")="无" then
  		JMail.Sender = "anonymous@of1884.org"
  	Else
		JMail.Sender = request.Form("smail")
	End if
	JMail.SenderName = getUserName(session("user_id"))
	JMail.AddRecipient request.Form("rmail")
	JMail.Subject = request.Form("subject")
	JMail.Body = request.Form("content")
	JMail.Priority = 3
	JMail.Execute
	JMail.Close
	Set JMail = Nothing
   
   	Response.write "<script language='javascript'>window.close()</script>"
End if
%>        
    </td>
  </tr>
</table>
<%
'## 关闭数据库连接
conn.close
set conn=nothing
%>
</body>
</html>


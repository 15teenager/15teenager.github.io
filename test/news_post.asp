<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout.asp"
end if

if getUserLevel(session("user_id"))<=1 then 
	Response.redirect("home.asp")
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

If Request.ServerVariables("Request_Method")="GET" then
   method=lcase(request.QueryString("method"))
   If  method="" then
	method="new"
   End if

   If method = "edit" then
	if session("news_id") <> "" then
		newsid = session("news_id")	
	else
		Response.Redirect "home.asp"
	end if
	
	'##     �����ݿ�����
	strSql = "SELECT N_Title, N_Content, N_Album_ID FROM News "
	strSql = strSql & "where News.News_ID = " &  newsid
	set rs = conn.Execute (StrSql)
	
	txtsub = rs("N_Title")
	txtmsg=rs("N_Content")
	albumid = rs("N_Album_ID")
	
	rs.close
	set rs=nothing
   End if
%>
<form action="news_post.asp" method="post">
<table width="760" border="0" cellspacing="2" cellpadding="0" align="center" style="border: 1px solid rgb(0,0,0)" bgcolor="<% =TableBgColor %>">
<tbody>
<tr>
<td>&nbsp;</td>
</tr>
<tr><td>
<table width="90%" border="0" cellspacing="0" cellpadding="3" align="center">
<tr> 
<td width="60" noWrap><font face="<% =DefaultFontFace %>" size="2"><b>����:</b></font></td>
<td>
<input maxLength="80" name="Title" size="80" value="<% =cleancode(txtsub) %>">
</td></tr>
<tr>
<td noWrap vAlign="top"><font face="<% =DefaultFontFace %>" size="2"><b>����:<br>*<a tabindex="-1" href="javascript:alert('                                      С���ɣ�����\r\r�������������м������·�������ǿ����Ч������������           \r�ԡ�^_^\r\r    [:)]��[:P]��[:(]��[;)]\r\r�������ϤHTML����Ҳ���������·�������ʾ������;����\r�Ŷ�Ӧ��ϵ���£�\r -------------------------------------------------------------------------\r      ����վ                        		       HTML\r -------------------------------------------------------------------------\r [b]��[/b]                             <b>��</b>\r [i]��[/i]                               <i>��</i>\r [a]��[/a]                            <a>��</a>\r [quote]��[/quote]             <BLOCKQUOTE>��</BLOCKQUOTE>\r [code]��[/code]                 <pre>��</pre>\r\r')"><font color='red' size="1">С����</font></a>*</b></font> 
</td>
<td><textarea cols="80" name="Message" rows="10" wrap="VIRTUAL"><%=cleancode(txtMsg)%></textarea><br></td></tr>
<tr>
<td noWrap><font face="<% =DefaultFontFace %>" size="2"><b>���Ӱ��:</b></font></td>
<td>
<% =DoDropDown("Album  where A_Personal = false order by A_Posted_Date Desc", "A_Name", "Album_ID", albumid, "Album_id", 3, 1) %>
</td></tr>
<p>
<input name="method" type="hidden" value="<% =method %>"> 
<tr> 
      <td noWrap colspan="2" height="40" valign="bottom"> 
        <div align="center">
        <table width="100%">
          <tr>
            <td width="45%">
                           <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:document.forms[0].submit()"><font size=2>ȷ ��</font></a></div>
                		</td>
              		      </tr>
            		   </table>
            </td>
            <td width="10%">&nbsp;</td>
            <td width="45%">
                           <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="left">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:history.back()"><font size=2>ȡ ��</font></a></div>
                		</td>
              		      </tr>
            		   </table>
            </td>
          </tr>
        </table>           		   
        </div>
      </td>
</tr></table>
</td>
</tr>
<tr>
<td>&nbsp;</td>
</tr>
</tbody></table>
</form>
<%Else
   title=trim(chkString(request.Form("title")))
   msg=trim(chkString(request.Form("message")))
   
   If  title="" or msg = "" then
	GO_Result "��������ݲ���Ϊ�գ�", false, "javascript:history.back()"
   else
	userid=Session("user_id")	
	method=lcase(request.Form("method"))
	
	if method = "new" then
		strSql = "insert into News (N_title, N_Content, N_Posted_By, N_Album_ID) Values ('"
		strSql = StrSql & title & "', '"
		strSql = StrSql & msg & "', "
		strSql = StrSql & userid & ", "
		strSql = StrSql & request.Form("Album_id") & ")"
		'Response.write strsql
		conn.Execute StrSql
		
		if Err.description <> "" then 
			GO_Result "�������ݴ��� " & Err.description, false, "javascript:history.back()"
		Else
			session("news_pageno")=1
			Go_Result  "���ŷ���ɹ�!", true, "home.asp"
		End If
	end if
	
	if method = "edit" then
		if session("news_id") <> "" then
			newsid = session("news_id")	
		else
			Response.Redirect "home.asp"
		end if
		
		strSql ="SELECT News.N_Posted_By from News where News_ID = " & newsid 
		'Response.Write StrSql
		set rs = conn.Execute (StrSql)

		If rs.Eof or rs.Bof then  
			GO_Result "���Ų�����!", false, "javascript:history.back()"
		Elseif rs("N_Posted_By")= userid or getUserLevel(userid)=3 then 
			msg = msg & vbcrlf & vbcrlf & " --- "& getUserName(userid) & " �޸���" & now()
			strSql = "update News set N_Title = '" & title
			strSql = StrSql & "', N_Content = '" & msg
			strSql = StrSql & "', N_Album_ID = " & request.Form("Album_id")
			strSql = StrSql & " where News_ID=" & newsid
			'Response.write strsql
			conn.Execute strSql
			
			if Err.description <> "" then 
				GO_Result "�޸Ĵ��� " & Err.description, false, "javascript:history.back()"
			Else
				Go_Result  "�����޸ĳɹ���", true, "news_info.asp"
			End If
		Else	
			GO_Result "��û��Ȩ���޸�����!", false, "javascript:history.back()"
		End if
		
		rs.close
		set rs=nothing
	end if 
   end if
End if

'## �ر����ݿ�����
conn.close
set conn=nothing
%>	
</body>
</html>

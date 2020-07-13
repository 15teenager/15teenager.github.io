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
	
	'##     打开数据库连接
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
<td width="60" noWrap><font face="<% =DefaultFontFace %>" size="2"><b>标题:</b></font></td>
<td>
<input maxLength="80" name="Title" size="80" value="<% =cleancode(txtsub) %>">
</td></tr>
<tr>
<td noWrap vAlign="top"><font face="<% =DefaultFontFace %>" size="2"><b>内容:<br>*<a tabindex="-1" href="javascript:alert('                                      小技巧！！！\r\r你可以在输入框中加上以下符号以增强感情效果，不信你试           \r试。^_^\r\r    [:)]、[:P]、[:(]、[;)]\r\r如果你熟悉HTML，你也可以用以下符号来表示特殊用途。符\r号对应关系如下：\r -------------------------------------------------------------------------\r      本网站                        		       HTML\r -------------------------------------------------------------------------\r [b]，[/b]                             <b>，</b>\r [i]，[/i]                               <i>，</i>\r [a]，[/a]                            <a>，</a>\r [quote]，[/quote]             <BLOCKQUOTE>，</BLOCKQUOTE>\r [code]，[/code]                 <pre>，</pre>\r\r')"><font color='red' size="1">小技巧</font></a>*</b></font> 
</td>
<td><textarea cols="80" name="Message" rows="10" wrap="VIRTUAL"><%=cleancode(txtMsg)%></textarea><br></td></tr>
<tr>
<td noWrap><font face="<% =DefaultFontFace %>" size="2"><b>相关影集:</b></font></td>
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
                  		  <div align="center" class="p9"><A href="javascript:document.forms[0].submit()"><font size=2>确 认</font></a></div>
                		</td>
              		      </tr>
            		   </table>
            </td>
            <td width="10%">&nbsp;</td>
            <td width="45%">
                           <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="left">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:history.back()"><font size=2>取 消</font></a></div>
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
	GO_Result "标题和内容不能为空！", false, "javascript:history.back()"
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
			GO_Result "增加数据错误： " & Err.description, false, "javascript:history.back()"
		Else
			session("news_pageno")=1
			Go_Result  "新闻发表成功!", true, "home.asp"
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
			GO_Result "新闻不存在!", false, "javascript:history.back()"
		Elseif rs("N_Posted_By")= userid or getUserLevel(userid)=3 then 
			msg = msg & vbcrlf & vbcrlf & " --- "& getUserName(userid) & " 修改于" & now()
			strSql = "update News set N_Title = '" & title
			strSql = StrSql & "', N_Content = '" & msg
			strSql = StrSql & "', N_Album_ID = " & request.Form("Album_id")
			strSql = StrSql & " where News_ID=" & newsid
			'Response.write strsql
			conn.Execute strSql
			
			if Err.description <> "" then 
				GO_Result "修改错误： " & Err.description, false, "javascript:history.back()"
			Else
				Go_Result  "新闻修改成功！", true, "news_info.asp"
			End If
		Else	
			GO_Result "你没有权限修改新闻!", false, "javascript:history.back()"
		End if
		
		rs.close
		set rs=nothing
	end if 
   end if
End if

'## 关闭数据库连接
conn.close
set conn=nothing
%>	
</body>
</html>

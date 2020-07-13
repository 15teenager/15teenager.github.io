<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout.asp"
end if
%>

<html>
<head>
<title>影集</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
<script language="JavaScript">
<!--
function chkdirectory() {
	if (document.forms[0].aName.value=="") {
		alert("影集名称不能为空！");
		return;
	}
	if (document.forms[0].directory.value=="") {
		alert("影集目录不能为空！");
		return;
	}
	if (validateFileEntry(document.forms[0].directory.value)==true) {
		document.forms[0].submit();
	}
}

function validateFileEntry(validString) {
	var isCharValid = true;
	var inValidChar;
	for (i=0 ; i < validString.length ; i++) {
		if (validateCharacter(validString.charAt(i)) == false) {
			isCharValid = false;
			inValidChar = validString.charAt(i);
			i=validString.length;
		}
	}	   
	if (i < 1) { return false; }	   
	if (isCharValid == false) {
		if (inValidChar) { alert("无效的目录名. 不能含有 '" + inValidChar + "'.");	}
		else             { alert("无效的目录名. 请重新输入."); }
		return false;
   }
   return true;
}
	
function validateCharacter(character) {
   if ((character >= 'a' && character <= 'z') || ( character >='A' && character <='Z') || ( character >= '0' && character <= '9') || ( character =='-') || ( character == '_')) return true;	
   else return false;
}

// -->
</script>
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
   	albumid=session("album_id")
	
	'##     打开数据库连接
	strSql = "SELECT * FROM Album where Album_ID = " &  albumid
	set rs = conn.Execute (StrSql)
	
	txtsub = rs("A_Name")
	txtmsg=rs("A_Comment")
	txtdir =rs("A_Directory")
	
	rs.close
	set rs=nothing
	
	Response.Write "<br>"
   End if
%>   
<form action="album_post.asp" method="post">
<table width="760" border="0" cellspacing="2" cellpadding="0" align="center" style="border: 1px solid rgb(0,0,0)" bgcolor="<% =TableBgColor %>">
<tbody>
<tr>
<td>&nbsp;</td>
</tr>
<tr><td>
<table width="90%" border="0" cellspacing="0" cellpadding="3" align="center">
<tr> 
      <td width="60" noWrap><font face="<% =DefaultFontFace %>" size="2"><b>名称:</b></font></td>
<td>
        <input maxLength="80" name="aName" size="80" value="<% =cleancode(txtsub) %>">
</td></tr>
<tr>
      <td noWrap vAlign="top"><font face="<% =DefaultFontFace %>" size="2"><b>说明:<br>
        *<a tabindex="-1" href="javascript:alert('                                      小技巧！！！\r\r你可以在输入框中加上以下符号以增强感情效果，不信你试           \r试。^_^\r\r    [:)]、[:P]、[:(]、[;)]\r\r如果你熟悉HTML，你也可以用以下符号来表示特殊用途。符\r号对应关系如下：\r -------------------------------------------------------------------------\r      本网站                        		       HTML\r -------------------------------------------------------------------------\r [b]，[/b]                             <b>，</b>\r [i]，[/i]                               <i>，</i>\r [a]，[/a]                            <a>，</a>\r [quote]，[/quote]             <BLOCKQUOTE>，</BLOCKQUOTE>\r [code]，[/code]                 <pre>，</pre>\r\r')"><font color='red' size="1">小技巧</font></a>*</b></font> 
</td>
<td>
        <textarea cols="80" name="comment" rows="10" wrap="VIRTUAL"><%=cleancode(txtmsg)%></textarea>
        <br></td></tr>
<tr>
      <td noWrap><font face="<% =DefaultFontFace %>" size="2"><b>目录:</b></font></td>
      <td>
      <%If method = "edit" then%>
        <input type="text" name="directory" size="30" maxlength="30" value="<% =txtdir %>" disabled>
      <%Else%>
        <input type="text" name="directory" size="30" maxlength="30">
      <%End if%>
      </td>
    </tr>
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
                  		  <div align="center" class="p9"><A href="javascript:chkdirectory()"><font size=2>确 认</font></a></div>
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
	title=trim(chkString(request.Form("aName")))
	msg=trim(chkString(request.Form("comment")))
	
	method=lcase(request.Form("method"))		
	if method = "new" then
		dir=trim(request.Form("directory"))
		
		strSql ="SELECT * FROM Album where A_Directory = '" & dir & "'" 
		set rs = conn.Execute (strSql)
		if rs.Eof or rs.Bof then
			strSql = "insert into Album (A_Name, A_Comment, A_Directory) Values ('"
			strSql = StrSql & title & "', '"
			strSql = StrSql & msg & "', '"
			strSql = StrSql & dir & "')"
			'Response.write strSql
			conn.Execute strSql
		
			if Err.description <> "" then 
				GO_Result "数据库错误： " & Err.description, false, "javascript:history.back()"
			Else
				Go_Result  "影集创建成功！", true, "album.asp"
			End If
		else
			GO_Result "目录“" & dir & "”已经存在，请重新输入。", false, "javascript:history.back()"
		end if
		
		rs.close
		set rs=nothing
	end if
	
	if method = "edit" then
		if session("album_id") <> "" then
			albumid=session("album_id")	
		else
			Response.Redirect "album_info.asp"
		end if
		
		strSql = "SELECT * FROM Album where Album_ID = " &  albumid
		'Response.Write StrSql
		set rs = conn.Execute (StrSql)

		If rs.Eof or rs.Bof then  
			GO_Result "影集不存在!", false, "javascript:history.back()"
		Elseif chkPermission(albumid)=true then 
			msg = msg & vbcrlf & vbcrlf & " --- "& getUserName(Session("user_id")) & " 修改于" & now()
			strSql = "update Album set A_Name = '" & title
			strSql = StrSql & "', A_Comment = '" & msg
			strSql = StrSql & "', A_Posted_Date=#" & now() 
			strSql = strSql & "# where Album_ID=" & albumid
			'Response.write strsql
			conn.Execute strSql
			
			if Err.description <> "" then 
				GO_Result "数据库错误： " & Err.description, false, "javascript:history.back()"
			Else
				Go_Result  "影集修改成功！", true, "album_info.asp"
			End If
			
			'刷新用户时间
			ownerid=getAlbumOwner(albumid)
			If ownerid<>0 then  
				strSql = "update UserInfo set U_Updated_Date=#" & now() & "# where User_ID=" & ownerid
				'Response.write strsql
				conn.Execute strSql		
			End if
		Else	
			GO_Result "你没有权限修改这个影集！", false, "javascript:history.back()"
		End if
		
		rs.close
		set rs=nothing
	end if	
End if

'## 关闭数据库连接
conn.close
set conn=nothing
%>
</body>
</html>
	
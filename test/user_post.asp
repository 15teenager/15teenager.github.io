<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout2.asp"
end if
%>

<html>
<head>
<title>用户管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
<script language="JavaScript">
<!--
function chkdirectory() {
	if (document.forms[0].uName.value=="") {
		alert("用户名不能为空！");
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
<body><br>
<%
if getUserLevel(session("user_id"))<3 then
	GO_Result "你没有权限进行用户管理!", false, "javascript:history.back()"
	Response.End
end if

set conn=server.createobject("adodb.connection")
conn.open ConnString
%>
<table width="450" height="310" align="center" border="0" cellspacing="2" cellpadding="0"> 
  <tr valign="top"> 
    <td>
<%If Request.ServerVariables("Request_Method")="GET" then
   method=lcase(request.QueryString("method"))

   If  method="" then
	method="new"
   End if

   If method = "edit" then
   	uid = request.QueryString("id")
	strSql ="SELECT * FROM Members INNER JOIN UserInfo ON Members.Member_id = UserInfo.User_ID where Members.Member_id =" & uid
	set rs = conn.Execute (StrSql)
	
	txtUname = rs("M_Name")
	txtLevel = rs("M_Level")
	isTeacher = rs("U_Is_Teacher")
	albumid = rs("U_Album_ID")
	
	rs.close
	set rs=nothing
	
	'##     打开数据库连接
	strSql = "SELECT * FROM Album where Album_ID = " &  albumid
	set rs = conn.Execute (StrSql)
	
	txtdir =rs("A_Directory")
	
	rs.close
	set rs=nothing
	
	Response.Write "<br>"
   End if
%>    
      <form method="post" action="user_post.asp">
        <table width="80%" border="0" align="center">
          <tr> 
            <th colspan="2"><font face="<% =SpecificFontFace %>">
            <% If method = "edit" then %>
            	修改用户
            <% Else %>
            	增加用户
            <% End if %>
            </font></th>
          </tr>
          <tr> 
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr> 
            <td height="30" width="25%">用户名：</td>
            <td height="30" width="81%"> 
              <input type="text" name="uName" value="<%=txtUname%>" size="8" maxlength="8">
            </td>
          </tr>
          <tr> 
            <td height="30" width="25%">级别：</td>
            <td height="30" width="81%"> 
              <% =DoDropDown("UserLevel", "L_Name", "Level_ID", txtLevel, "level", 0, 1) %>
            </td>
          </tr>
          <tr> 
            <td height="30" width="25%">类别：</td>
            <td height="30" width=81%> 
              <table width="100%" border="0">
                <tr>
                  <td width="34%"> 
                    <input type="radio" name="mtype" value="False" <%=Chked(Not isTeacher)%>>
                    学生 </td>
                  <td width="66%"> 
                    <input type="radio" name="mtype" value="True" <%=Chked(isTeacher)%>>
                    教师</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td width=25% height="30">影集目录：</td>
            <td width=81% valign="top" height="30"> 
            <% If method = "edit" then%>
            	<input type="text" name="directory" value="<%=txtdir%>" size="30" maxlength="30" disabled>
            <% Else %>
            	<input type="text" name="directory" value="<%=txtdir%>" size="30" maxlength="30">
            <% End if %>
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
                          <div align="center" class="p9"><A href="javascript:chkdirectory()"><font size=2>确 
                            认</font></a></div>
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td width="15%" height="58">&nbsp;</td>
                  <td width="45%" align="left" valign="middle" height="58"> 
                    <table border="1" cellspacing="0" cellpadding="2" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="left">
                      <tr> 
                        <td> 
                          <div align="center" class="p9"><A href="javascript:window.close()"><font size=2>关 
                            闭</font></a></div>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        <input name="id" type="hidden" value="<% =uid %>"> 
        <input name="method" type="hidden" value="<% =method %>"> 
        
      </form>
<%Else
	txtUname=request.Form("uName")
	txtLevel=request.Form("level")
	TF=Request.Form("mtype")
	txtdir=request.Form("directory")
	
	method=lcase(request.Form("method"))		
	if method = "new" then
		strSql ="SELECT * FROM Members where M_Name = '" & txtUname & "'" 
		set rs = conn.Execute (strSql)
		
		if rs.Eof or rs.Bof then
			strSql ="SELECT * FROM Album where A_Directory = '" & txtdir & "'" 
			set rs1 = conn.Execute (strSql)
			if rs1.Eof or rs1.Bof then
				strSql = "insert into Members (M_Name, M_Password, M_Level) Values ('"
				strSql = StrSql & txtUname & "', '"
				strSql = StrSql & SiteName & "', "
				strSql = StrSql & txtLevel & ")"
				'Response.write strSql
				conn.Execute strSql
				
				strSql ="SELECT * FROM Members where M_Name = '" & txtUname & "'" 
				set rs3 = conn.Execute (strSql)
				uid=rs3("Member_id")
				rs3.close
				set rs3=nothing
				
				strSql = "insert into Album (A_Name, A_Directory, A_Personal) Values ('"
				strSql = StrSql & txtUname & "', '"
				strSql = StrSql & txtdir & "', "
				strSql = StrSql & "true )"
				'Response.write strSql
				conn.Execute strSql
				
				strSql ="SELECT * FROM Album where A_Directory = '" & txtdir & "'" 
				set rs3 = conn.Execute (strSql)
				albumid=rs3("Album_ID")
				rs3.close
				set rs3=nothing
				
				strSql = "insert into UserInfo (User_ID, U_Updated_By, U_Is_Teacher, U_Album_ID) Values ("
				strSql = StrSql & uid & ", "
				strSql = StrSql & session("user_id") & ", "
				strSql = StrSql & TF & ", "
				strSql = StrSql & albumid & ")"
				'Response.write strSql
				conn.Execute strSql
				
				Response.write "<script language='javascript'>window.close()</script>"
			else
				GO_Result "目录“" & txtdir & "”已经存在，请重新输入。", false, "javascript:history.back()"
			end if
			
			rs1.close
			set rs1=nothing	
		else
			GO_Result "用户“" & txtUname & "”已经存在，请重新输入。", false, "javascript:history.back()"
		end if
		
		rs.close
		set rs=nothing
	end if
	
	if method = "edit" then
		uid=request.Form("id")
		
		strSql ="SELECT * FROM Members where Member_id = " & uid
		'Response.write strsql
		set rs = conn.Execute (strSql)

		If rs.Eof or rs.Bof then  
			GO_Result "用户不存在!", false, "javascript:history.back()"
		Else
			strSql = "update Members set M_Name = '" & txtUname
			strSql = StrSql & "', M_Level = " & txtLevel
			strSql = strSql & " where Member_ID=" & uid
			'Response.write strsql
			conn.Execute strSql
			
			strSql = "update UserInfo set U_Updated_By = " & session("user_id")
			strSql = StrSql & ", U_Updated_Date=#" & now() 
			strSql = StrSql & "#, U_Is_Teacher=" & TF
			strSql = strSql & " where User_ID=" & uid
			'Response.write strsql
			conn.Execute strSql
			
			Response.write "<script language='javascript'>window.close()</script>"
		End if
		
		rs.close
		set rs=nothing
	end if 	
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
	
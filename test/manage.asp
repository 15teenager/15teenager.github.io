<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout.asp"
end if
%>

<%
Function Chked(YN)
   '  To Check Check Boxes
   if YN = "true" then
      Chked = "Checked"
   else 
      Chked = ""
   end if    
End Function

Function IsTeacher(YN)
   '  To Check Check Boxes
   if YN or YN = "true" then
      IsTeacher = "教师"
   else 
      IsTeacher = "学生"
   end if    
End Function
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
<script language="JavaScript">
<!--
function WinOpen(url) {
	var AtWnd = window.open(url,"PopWindow","toolbar=no,directories=no,menubar=no,location=no,status=no,scrollbars=no,resizable=no,copyhistory=no,width=500,height=350");
	if ((document.AttachForm.window != null) && (!AtWnd.opener)) {
		AtWnd.opener = document.AttachForm.window;
		AtWnd.focus();
	}
}

function doConfirm()
{
  if (window.confirm("确实要删除吗？"))
  	document.forms[0].submit();
}
// -->
</script>
</head>
<body>
<%
if getUserLevel(session("user_id"))<3 then
	GO_Result "你没有权限进行管理!", false, "javascript:history.back()"
	Response.End
end if

set conn=server.createobject("adodb.connection")
conn.open ConnString

mode=request.QueryString("mode")
If  mode="" then
	if session("manage_mode") <> "" then
		mode = session("manage_mode")	
	else
   		mode=1
   	end if
end if
session("manage_mode")=mode

m=request.Form("m")
%>
<table width="760" border="0" align="center" height="320" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="118" valign="top" height="317"> 
      <table width="100%" border="0" height="300" cellspacing="0" cellpadding="0">
        <tr valign="top"> 
          <td height="320"> 
            <table width="100%" border="0">
              <tr bgcolor="#CCCCFF"> 
                <td height="25"><div align="center">
<% if mode=1 then%>
		登录信息
<%else%>
                <a href="manage.asp?mode=1">登录信息</a>
<%end if%>
		</div></td>
              </tr>
              <tr bgcolor="#CCCCFF"> 
                <td height="25"><div align="center">
<% if mode=2 then%>
		浏览记录
<%else%>
                <a href="manage.asp?mode=2">浏览记录</a>
<%end if%>
		</div></td>                
              </tr>
              <tr bgcolor="#CCCCFF"> 
                <td height="25"><div align="center">
<% if mode=3 then%>
		用户管理
<%else%>
                <a href="manage.asp?mode=3">用户管理</a>
<%end if%>
		</div></td>                
              </tr>  
              <tr bgcolor="#CCCCFF"> 
                <td height="25"><div align="center">
<% if mode=4 then%>
		数据备份
<%else%>
                <a href="manage.asp?mode=4">数据备份</a>
<%end if%>
		</div></td>                
              </tr>                          
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="640" valign="top" height="317" align="center"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="3" style="border: 1px solid rgb(0,0,0)" height="320" bgcolor="<% =TableBgColor %>">
        <tr valign="top" align="center"> 
          <td> <br>  
<%if m = "" then %>
            <form method="post" name="AttachForm" action="manage.asp">
              <table width="90%" border="0">
                <tr align="center" valign="top"> 
                  <td height="43"> 
<%	Select Case mode
	case 1   
		'##     打开数据库连接
		strSql ="SELECT * FROM Login order by L_Date"

		set rs=server.createobject("adodb.recordset")
		rs.cachesize=Pagesize
		rs.open strSql,conn,3
	
		mypage=request.QueryString("pageno")
		If  mypage="" then
			if session("manage_pageno1") <> "" then
				mypage = session("manage_pageno1")	
			else
   				mypage=1
   			end if
		end if
		session("manage_pageno1")=mypage
%>                    
                    <table width="100%" border="0" cellspacing="2">
                      <tr bgcolor="#9999FF"> 
                        <td width="10"> 
                          <div align="center">选择</div>
                        </td>
                        <td width="60"> 
                          <div align="center">姓名</div>
                        </td>
                        <td width="20%"> 
                          <div align="center">登录类型</div>
                        </td>
                        <td width="20%"> 
                          <div align="center">用户IP地址</div>
                        </td>
                        <td width="25%"> 
                          <div align="center">时间</div>
                        </td>
                      </tr>
<% 
		'显示所有login event
		If rs.Eof or rs.Bof then  
			' No items found in DB
			Response.Write "<tr><td>&nbsp;</td><td collspan=4>没有记录！</td></tr>"
		Else
			showBtn="true"
			rs.movefirst
	
			rs.pagesize=Pagesize
			maxpages=cint(rs.pagecount)
			if cint(mypage) > maxpages then 
				mypage=maxpages
				session("manage_pageno1")=mypage
			end if
			rs.absolutepage=mypage
	
			rec = 1
			do until rs.Eof or rec > Pagesize '## Display Forum
				if rs("L_Type") = "合法" then 
					CColor = AltForumCellColor
				else
					CColor = "#FF3C00"
				End if
	  
	  			Response.Write "<tr>"
		  		Response.Write "<td width='10' bgcolor='" & CColor & "' valign='top' align='center'><input type='checkbox' name='SelectedID' value='" & rs("Login_ID") & "'" & Chked(request.QueryString("a")) & "></td>"
		  		Response.Write "<td width='60' bgcolor='" & CColor & "' valign='top' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>" &  rs("L_Name") & "</font></td>"
	  			Response.Write "<td width='20%' bgcolor='" & CColor & "' valign='top' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>" & rs("L_Type") & "</font></td>"
	  			Response.Write "<td width='20%' bgcolor='" & CColor & "' valign='top' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>" & rs("L_IP") & "</font></td>"
	  			Response.Write "<td width='25%' bgcolor='" & CColor & "' valign='top' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='1'>" & rs("L_Date") & "</font></td>"
		  		Response.Write "</tr>"
		  		rs.MoveNext
	
			    	rec = rec + 1
			loop
		End If

		rs.close
		set rs=nothing
%>                      
                    </table>                   
<%	case 2  
		'##     打开数据库连接
		strSql ="SELECT * FROM Events order by L_Visit_Date"

		set rs=server.createobject("adodb.recordset")
		rs.cachesize=Pagesize
		rs.open strSql,conn,3
	
		mypage=request.QueryString("pageno")
		If  mypage="" then
			if session("manage_pageno2") <> "" then
				mypage = session("manage_pageno2")	
			else
   				mypage=1
   			end if
		end if
		session("manage_pageno2")=mypage	
%>                    
                    <table width="100%" border="0" cellspacing="2">
                      <tr bgcolor="#9999FF"> 
                        <td width="10"> 
                          <div align="center">选择</div>
                        </td>
                        <td width="20%"> 
                          <div align="center">姓名</div>
                        </td>
                        <td width="15%"> 
                          <div align="center">事件类型</div>
                        </td>
                        <td width="20%"> 
                          <div align="center">当前页面</div>
                        </td>
                        <td width="20%"> 
                          <div align="center">前一页面</div>
                        </td>
                        <td width="30%"> 
                          <div align="center">时间</div>
                        </td>
                      </tr>
<% 
		'显示所有event
		If rs.Eof or rs.Bof then  
			' No items found in DB
			Response.Write "<tr><td>&nbsp;</td><td collspan=5>没有记录！</td></tr>"
		Else
			showBtn="true"
			rs.movefirst
	
			rs.pagesize=Pagesize
			maxpages=cint(rs.pagecount)
			if cint(mypage) > maxpages then 
				mypage=maxpages
				session("manage_pageno2")=mypage
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
	  			Response.Write "<td width='10' bgcolor='" & CColor & "' align='center'><input type='checkbox' name='SelectedID' value='" & rs("Log_ID") & "'" & Chked(request.QueryString("a")) & "></td>"
	  			Response.Write "<td width='20%' bgcolor='" & CColor & "' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>" &  getUserName(rs("L_Visited_By")) & "</font></td>"
		  		Response.Write "<td width='15%' bgcolor='" & CColor & "' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>" &  rs("L_Event") & "</font></td>"
		  		Response.Write "<td width='20%' bgcolor='" & CColor & "' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'><a href='" & rs("L_ScriptName") & "?" & rs("L_QueryString") & "' target='_blank'>" & rs("L_ScriptName") & "?" & rs("L_QueryString") & "</a></font></td>"
	  			Response.Write "<td width='20%' bgcolor='" & CColor & "' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'><a href='" & rs("L_Referer") & "' target='_blank'>" & rs("L_Referer") & "</a></font></td>"
	  			Response.Write "<td width='30%' bgcolor='" & CColor & "' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='1'>" & rs("L_Visit_Date") & "</font></td>"
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
<%	case 3   
		'##     打开数据库连接
		strSql ="SELECT * FROM Members INNER JOIN UserInfo ON Members.Member_id = UserInfo.User_ID order by Members.M_Name"

		set rs=server.createobject("adodb.recordset")
		rs.cachesize=Pagesize
		rs.open strSql,conn,3
	
		mypage=request.QueryString("pageno")
		If  mypage="" then
			if session("manage_pageno3") <> "" then
				mypage = session("manage_pageno3")	
			else
   				mypage=1
   			end if
		end if
		session("manage_pageno3")=mypage
%>                   
                    <table width="100%" border="0" cellspacing="2">
                      <tr bgcolor="#9999FF"> 
                        <td width="5%"> 
                          <div align="center">选择</div>
                        </td>
                        <td width="15%"> 
                          <div align="center">姓名</div>
                        </td>
                        <td width="20%"> 
                          <div align="center">级别</div>
                        </td>
                        <td width="15%"> 
                          <div align="center">类别</div>
                        </td>
                        <td width="30%"> 
                          <div align="center">影集目录</div>
                        </td>
                      </tr>
<% 
		'显示所有用户
		If rs.Eof or rs.Bof then  
			' No items found in DB
			Response.Write "<tr><td>&nbsp;</td><td collspan=4>没有记录！</td></tr>"
		Else
			showBtn="true"
			rs.movefirst
	
			rs.pagesize=Pagesize
			maxpages=cint(rs.pagecount)
			if cint(mypage) > maxpages then 
				mypage=maxpages
				session("manage_pageno3")=mypage
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
				
				strSql ="SELECT Album.A_Name, Album.A_Directory FROM Album where Album.Album_ID = " & rs("U_Album_ID")
				'Response.write strSql
				set rs1 = conn.Execute (strSql)

	  			Response.Write "<tr>"
		  		Response.Write "<td width='5%' bgcolor='" & CColor & "' valign='top' align='center'><input type='radio' name='SelectedID' value='" & rs("Member_id") & "'></td>"
		  		Response.Write "<td width='15%' bgcolor='" & CColor & "' valign='top' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'><A href='#' onClick='WinOpen(""user_post.asp?method=edit&id=" & rs("Member_id") & """)'>" &  rs("M_Name") & "</a></font></td>"
	  			Response.Write "<td width='20%' bgcolor='" & CColor & "' valign='top' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>" & getLevelName(rs("M_Level")) & "</font></td>"
	  			Response.Write "<td width='15%' bgcolor='" & CColor & "' valign='top' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>" & IsTeacher(rs("U_Is_Teacher")) & "</font></td>"
	  			Response.Write "<td width='30%' bgcolor='" & CColor & "' valign='top' align='left'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>" & rs1("A_Directory") & "</font></td>"
		  		Response.Write "</tr>"
		  		rs.MoveNext
				
				rs1.close
				set rs1=nothing
				
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
<%	case 4
		strSql = "SELECT * FROM Backup"
		set rs = conn.Execute (StrSql)
%>                   
		   <table width="100%" border="0" cellspacing="2">
                      <tr> 
                        <td height="15"> 
                          &nbsp;
                        </td>
                      </tr>
                      <tr> 
                        <td height="60" valign="top"> 
                          <div align="center">最近一次数据备份由<font face="<% =SpecificFontFace %>" color="<% =SpecificFontColor %>"><% =getUserName(rs("B_Backup_By"))%></font>备份于<font face="<% =SpecificFontFace %>" color="<% =SpecificFontColor %>"><% =formatDate(rs("B_Last_Backup")) %></font></div>
                        </td>
                      </tr>
                      <tr> 
                        <td> 
                          <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:document.forms[0].submit()"><font size=2>备 份</font></a></div>
                		</td>
              		      </tr>
            		   </table>
                        </td>
                      </tr>
                   </table>                                        
<%	End Select%>      		
 
<%	if maxpages > 1 then  %>
		<hr><div align=left><font face="<% =DefaultFontFace %>" size="2">
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
<%	End if%>              
                  <br></td>
                </tr>
<% if showBtn="true" then %>               
                <tr align="center"> 
                  <td> 
                    <table width="100%" border="0">
                      <tr> 
                        <td width="46%"> 
                          <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9">
                  		  <% if mode=3 then %>
                  		  	<A href='#' onClick='WinOpen("user_post.asp")'><font size=2>增 加</font></a>
                  		  <% else %>
                  		  	<a href="manage.asp?a=true"><font size=2>全部选中</font></a>
                  		  <% end if %></div>
                		</td>
              		      </tr>
            		   </table>
                        </td>
                        <td width="7%">&nbsp;</td>
                        <td width="47%"> 
                           <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="left">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:doConfirm()"><font size=2>删 除</font></a></div>
                		</td>
              		      </tr>
            		   </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
<% end if %>                
              </table>
              <input name="m" type="hidden" value="<% =mode %>">
            </form>
<%else
	Select Case m
		case 1
			if Request.Form("SelectedID")<>"" then
				For Each id in Request.Form("SelectedID")
					strSQL = "Delete * From Login where login_ID = " & id
					conn.Execute strSQL
				Next
				GO_Result "成功删除记录!", true, "manage.asp"
			else
				GO_Result "没有选择要删除的记录!", false, "javascript:history.back()"
			end if
		case 2
			if Request.Form("SelectedID")<>"" then
				For Each id in Request.Form("SelectedID")
					strSQL = "Delete * From Events where Log_ID = " & id
					conn.Execute strSQL
				Next
				GO_Result "成功删除记录!", true, "manage.asp"
			else
				GO_Result "没有选择要删除的记录!", false, "javascript:history.back()"
			end if
		case 3
			if Request.Form("SelectedID")<>"" then
				uid = Request.Form("SelectedID")
				
				if cint(uid)=cint(session("user_id")) then
					GO_Result "不能删除当前用户！", false, "javascript:history.back()"
				else
				
				albumid = getAlbumID(uid)
			
				strSql ="SELECT * FROM Album where Album_id = " & albumid
				'Response.write strSql
				set rs = conn.Execute (strSql) 
				dir=Server.URLEncode(rs("A_Directory"))
				rs.close
				set rs=nothing
			
				strSQL = "Delete * From Photoes where P_Album_ID = " & albumid
				conn.Execute (StrSql)
			
				strSQL = "Delete * From Album where Album_ID = " & albumid
				conn.Execute (StrSql)
			
				strSQL = "Delete * From UserInfo where User_ID = " & uid
				conn.Execute (StrSql)
			
				strSQL = "Delete * From Members where Member_id = " & uid
				conn.Execute (StrSql)
			
				Response.redirect ("del.cgi?path=" & dir & "&location=" & Server.URLEncode("/manage.asp"))
				'Response.write "del.cgi?path=" & dir & "&location=" & Server.URLEncode("/manage.asp")
				end if
			else
				GO_Result "没有选择要删除的记录!", false, "javascript:history.back()"
			end if
		case 4
			strSql = "update Backup set B_Backup_By = " & session("user_id") & ", B_Last_Backup = #" & now() & "#"
			conn.Execute (StrSql)
			
			Response.redirect "backup.cgi"
	End Select
end if%>  
	    <br>        
          </td>
        </tr>
      </table>
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

	
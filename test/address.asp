<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout.asp"
end if
%>

<%
Function FormatString(str, link)
	str = Replace(str, "��", "<br>")
	str = Replace(str, ",", "<br>")
	str = Replace(str, ";", "<br>")
	
	if link = "true" then
		tx = split(str, "<br>")
		str = ""
			
		dim i
		for i = 0 to ubound(tx)
			str = str & "<a href='mailto:" & trim(tx(i)) & "'>" & tx(i) & "</a>"
					
			if i<ubound(tx) then
				str = str & "<br>"
			end if
		next
	end if
	
	FormatString=str
End Function

Function SplitCity(cid)
SplitCity = ""
if cid <> "" then
	Str2 = "SELECT * FROM City where City_ID = " & cid
	set rs_city = conn.Execute (str2)
	
	if rs_city.Eof or rs_city.Bof then
		mcity = "����"
	else
		mcity = rs_city("C_Name")
	end if
	
	rs_city.close
	set rs_city=nothing
	
	for i = 1 to len(mcity)-1
		SplitCity=SplitCity & mid(mcity,i,1) & "<br>"
	next
	SplitCity=SplitCity & mid(mcity,len(mcity))
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

mode=request.QueryString("mode")
If  mode="" then
	if session("address_mode") <> "" then
		mode = session("address_mode")	
	else
   		mode=1
   	end if
end if

profileid=request.QueryString("id")
if profileid<>"" and mode=3 then
	mode=1
else
	session("address_mode")=mode
end if

if mode=3 then
	profileid=session("user_id")
end if
%>
<table width="760" border="0" align="center" height="320" cellspacing="2" cellpadding="0">
  <tr>
    <td width="118" valign="top" height="317"> 
      <table width="100%" border="0" height="300" cellspacing="0" cellpadding="0">
        <tr valign="top"> 
          <td height="320"> 
            <table width="100%" border="0">
              <tr bgcolor="#CCCCFF"> 
                <td height="25"><div align="center">
<% if mode=1 then%>
		������ʾ�û�
<%else%>
                <a href="address.asp?mode=1">������ʾ�û�</a>
<%end if%>
		</div></td>
              </tr>
              <tr bgcolor="#CCCCFF"> 
                <td height="25"><div align="center">
<% if mode=2 then%>
		��ʾ�����û�
<%else%>
                <a href="address.asp?mode=2">��ʾ�����û�</a>
<%end if%>
		</div></td>                
              </tr>
              <tr bgcolor="#CCCCFF"> 
                <td height="25"><div align="center">
<% if mode=3 then%>
		��ʾ��������
<%else%>
                <a href="address.asp?mode=3">��ʾ��������</a>
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
          <td><br>
<%
if profileid<>"" then   '��ʾ��������
	strSQL  = " SELECT * FROM Members INNER JOIN UserInfo ON Members.Member_id = UserInfo.User_ID where Members.Member_id = " & profileid
	set rs = conn.Execute (strSQL)

	If rs.Eof or rs.Bof then  
		' No items found in DB
		Response.Write "û�з����������û���"
	Else
		ucom = rs("U_Comment")
%>
	    <table width="95%" border="0" cellspacing="0" cellpadding="3"  height="320">
              <tr> 
                <td width="1%" height="50">&nbsp;</td>
                <td width="14%" height="50" valign="top"><b><font face="<%=SpecificFontFace%>" size="5"><%=rs("M_Name")%></font></b></td>
                <td width="35%">&nbsp; </td>
                <td width="50%" colspan="2" valign="top">
                <Form name="AttachForm">
                  <table width=80 align="right"> 
                    <tr>
                      <% if cint(profileid) = cint(session("user_id")) or getUserLevel(session("user_id"))=3 then  %>
                      <td> 
      			<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        		  <tr> 
          		    <td> 
            		      <div align="center" class="p9"><a href='profile.asp?id=<%=rs("Member_id")%>'><font size=2>�޸�����</font></a></div>
          		    </td>
        		  </tr>
      			</table>
    		      </td>
    		      <% end if %>
    		      <td> 
      			<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        		  <tr> 
          		    <td> 
            		      <div align="center" class="p9"><a href='#' onClick='WinOpen("message.asp?mode=1&id=<%=rs("Member_id")%>")'><font size=2>���ҷ�ѶϢ</font></a></div>
          		    </td>
        		  </tr>
      			</table>
    		      </td>
    		      <% if getEmail(profileid)<>""  then %>
    		      <td> 
      			<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        		  <tr> 
          		    <td> 
            		      <div align="center" class="p9"><a href='#' onClick='WinOpen("mail.asp?id=<%=rs("Member_id")%>")'><font size=2>���ҷ��ʼ�</font></a></div>
          		    </td>
        		  </tr>
      			</table>
    		      </td>    		     
    		      <%end if%> 
    		      <td> 
      			<table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        		  <tr> 
          		    <td> 
            		      <div align="center" class="p9"><a href='#' onClick='WinOpen2("album_info.asp?album_id=<%=rs("U_Album_ID")%>")'><font size=2>�ҵ�Ӱ��</font></a></div>
          		    </td>
        		  </tr>
      			</table>
    		      </td>
    		    </tr>
    		  </table>
    		  </Form>
                </td>
              </tr>
              <tr> 
                <td width="1%" height="35">&nbsp;</td>
                <td width="14%" height="35" valign="top">�绰��</td>
                <td colspan="3" valign="top"><font color="<%=SpecificFontColor%>"><% if rs("U_Phone")="" then Response.write "<��>" else Response.write rs("U_Phone") end if%></font></td>
              </tr>
              <tr> 
                <td width="1%" height="35">&nbsp;</td>
                <td width="14%" height="35" valign="top">E-mail�� </td>
                <td colspan="3" valign="top"><font color="<%=SpecificFontColor%>"><% if rs("U_Email")="" then Response.write "<��>" else Response.write "<a href='mailto:" & rs("U_Email") & "'>" & rs("U_Email") & "</a>"  end if%></font></td>
              </tr>
              <tr> 
                <td width="1%" height="35">&nbsp;</td>
                <td width="14%" height="35">OICQ/ICQ��</td>
                <td height="35" colspan="2"><font color="<%=SpecificFontColor%>"><% if rs("U_Oicq")="" then Response.write "<��>" else Response.write rs("U_Oicq") end if%></font></td>
                <td width="30%">�û�����<font face="<%=SpecificFontFace%>" color="<%=SpecificFontColor%>"><%=getLevelName(rs("M_Level"))%></font></td>
              </tr>
              <tr> 
                <td width="1%" height="35">&nbsp;</td>
                <td width="14%" height="35">������ҳ��</td>
                <td colspan="2"><font color="<%=SpecificFontColor%>">
                <% url=rs("U_Homepage")
                if url ="" then 
                	Response.write "<��>"
                else
                	if lcase(left(url, 5)) = "http:" then
                		Response.write "<A href='" & url & "' Target=_Blank>"
                	else
                		Response.write "<A href='http://" & url & "' Target=_Blank>"
                	end if
                	Response.write url & "</a>" 
                end if %></font></td>
                <td>���ڳ��У�<font face="<%=SpecificFontFace%>" color="<%=SpecificFontColor%>">
                <% city=getCity(rs("U_City"),rs("U_City2"))
                   if city="" then Response.write "<��>" else Response.write city end if%></font></td>
              </tr>
              <tr> 
                <td width="1%" height="35">&nbsp;</td>
                <td width="14%" height="35">ͨѶ��ַ��</td>
                <td colspan="2"><font face="<%=SpecificFontFace%>" color="<%=SpecificFontColor%>"><% if rs("U_Address")="" then Response.write "<��>" else Response.write rs("U_Address") end if%></font></td>
                <td>�������룺<font color="<%=SpecificFontColor%>"><% if rs("U_Zipcode")="" then Response.write "<��>" else Response.write rs("U_Zipcode") end if%></font></td>
              </tr>
              <tr> 
                <td width="1%" height="35">&nbsp;</td>
                <td width="14%" height="35">���˼�飺</td>
                <td colspan="3">&nbsp; </td>
              </tr>
              <tr> 
                <td width="1%">&nbsp;</td>
                <td colspan="4" valign="top" ><font face="<%=SpecificFontFace%>"  color="<%=SpecificFontColor%>"><% if ucom= "" then  Response.Write "<��>"  else  Response.Write formatStr(ucom)  end if %></font><p></p></td>
              </tr>
              <tr> 
                <td width="1%" height="40">&nbsp;</td>
                <td colspan="4" height="40">�û����ʴ�����<font face="<%=SpecificFontFace%>"><%=rs("M_Visited_Times")%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������ʱ�䣺<font face="<%=SpecificFontFace%>"><%=formatDate(rs("M_Last_Visited"))%></font></td>
              </tr>  
              <tr> 
                <td width="1%" height="20">&nbsp;</td>
                <td colspan="4" height="20">*** ��������<b><font face="<%=SpecificFontFace%>"><%=getUserName(rs("U_Updated_By"))%></font></b>������<%=formatDate(rs("U_Updated_Date"))%> ***
                </td>
              </tr>
            </table>
<%
	End if
	
	rs.close
	set rs=nothing
	
	if mode<>3 then
%>
<br>
<div align="right">
  <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
    <tr> 
      <td> 
        <div align="center" class="p9"><a href="javascript:history.back()"><font size=2>����</font></a></div>
      </td>
    </tr>
  </table>
</div>
<%
	end if
Else
	isTeacher=request.QueryString("t")
	cityid=request.QueryString("c")
	
	if isTeacher="true" or cityid<>"" or mode=2 then  '��ʾ����
		strSql = ""
		if isTeacher="true" then
			strSql=" where UserInfo.U_Is_Teacher = true"
		end if

		if cityid<>"" then 	
			strSql=" where UserInfo.U_Is_Teacher = false and UserInfo.U_City = " & cityid
		end if
		
		strSql  = " SELECT * FROM Members INNER JOIN UserInfo ON Members.Member_id = UserInfo.User_ID " & strSql & " order by Members.M_Name"
		set rs = conn.Execute (strSql)
%>
            <table width="90%" border="0" cellspacing="2" cellpadding="0">
              <tr> 
              <td align="center" bgcolor="<% =HeadCellColor %>">&nbsp;</td>
              <td align="center" bgcolor="<% =HeadCellColor %>"><strong><font face="<% =DefaultFontFace %>" size="2" color="<% =HeadFontColor %>">����</font></strong></td>
              <td align="center" bgcolor="<% =HeadCellColor %>"><strong><font face="<% =DefaultFontFace %>" size="2" color="<% =HeadFontColor %>">�绰</font></strong></td>
              <td align="center" bgcolor="<% =HeadCellColor %>"><strong><font face="<% =DefaultFontFace %>" size="2" color="<% =HeadFontColor %>">E-mail</font></strong></td>
              <td align="center" bgcolor="<% =HeadCellColor %>"><strong><font face="<% =DefaultFontFace %>" size="2" color="<% =HeadFontColor %>">OICQ/ICQ</font></strong></td>
              <td align="center" bgcolor="<% =HeadCellColor %>"><strong><font face="<% =DefaultFontFace %>" size="2" color="<% =HeadFontColor %>">����</font></strong></td>
              </tr>
<% 
		If rs.Eof or rs.Bof then  
			' No items found in DB
			Response.Write "<tr><td>&nbsp;</td><td collspan=4>û�з����������û���</td></tr>"
		Else   '��ʾ�û�����
			rs.movefirst
	
			i=0
			do until rs.Eof  
				if i mod 2= 0 then 
					CColor = AltForumCellColor
				else
					CColor = ForumCellColor
				End if
	  
	  			Response.Write "<tr>"
	  			Response.Write "<td bgcolor='" & CColor & "' valign='middle' align='center'>" & isNew(rs("U_Updated_Date"),2) & "</td>" & vbcrlf
	  			Response.Write "<td width='60' bgcolor='" & CColor & "' valign='middle' align='left'><font face='" & DefaultFontFace & "' size='2'><a href='address.asp?id=" & rs("Member_id") & "'><img src='gif/profile.gif' border=0>" & rs("M_Name") & "</a>&nbsp;</font></td>"
	  			Response.Write "<td bgcolor='" & CColor & "' valign='middle' align='left'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>"
	  			if rs("U_Phone")="" then 
	  				Response.write "<��>" 
	  			else 
	  				Response.write FormatString(rs("U_Phone"), "false")
	  			end if
	  			Response.Write "</font></td>"
	  			Response.Write "<td bgcolor='" & CColor & "' valign='middle' align='left'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>" 
	  			if rs("U_Email")="" then
	  				Response.write "<��>" 
	  			else 
	  				Response.write FormatString(rs("U_Email"), "true")  
	  			end if
	  			Response.write "</font></td>"
	  			Response.Write "<td bgcolor='" & CColor & "' valign='middle' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>"
	  			if rs("U_Oicq")="" then 
	  				Response.write "<��>" 
	  			else 
	  				Response.write rs("U_Oicq") 
	  			end if
	  			Response.Write "</font></td>"
	  			Response.Write "<td width='40' bgcolor='" & CColor & "' valign='middle' align='center'><font face='" & DefaultFontFace & "' color='" & ForumFontColor & "' size='2'>" & getCity(rs("U_City"),rs("U_City2")) & "</font></td>"
	  			Response.Write "</tr>"
	  			rs.MoveNext
	  	
	    			i = i + 1
			loop
		End If  
		
		rs.close
		set rs=nothing    
%>
            </table>
<%		if mode<>2 then%>
<br>
<div align="right">
  <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
    <tr> 
      <td> 
        <div align="center" class="p9"><a href="javascript:history.back()"><font size=2>����</font></a></div>
      </td>
    </tr>
  </table>
</div>
<%
		end if            
	Else  'mode=1 %>
	    <div align="center"><b><font face="<%=SpecificFontFace%>" size="6"><% =PageTitle %></font></b> 
              <br>
              <hr width="90%">
              <br>
              <table width="90%" border="1" cellspacing="8" cellpadding="2" bordercolorlight="#78592c" bordercolordark="#faf5ed" bgcolor="#CCCCFF" height="178">
                <tr> 
                  <td height="50"> 
                    <p align="center"><font face="<%=SpecificFontFace%>" size="5"><a href="address.asp?t=true">��<br>
                      &nbsp;<br>
                      ʦ</a></font></p>
                  </td>	
                  <td height="50"> 
                    <p align="center"><font face="<%=SpecificFontFace%>" size="5"><a href="address.asp?c=1"><%=SplitCity(1)%><br>
                      ��<br>
                      ��</a></font></p>
                  </td>
                  <td height="50"> 
                    <div align="center"><font face="<%=SpecificFontFace%>" size="5"><a href="address.asp?c=2"><%=SplitCity(2)%><br>
                      ��<br>
                      ��</a></font></div>
                  </td>
                  <td height="50"> 
                    <div align="center"><font face="<%=SpecificFontFace%>" size="5"><a href="address.asp?c=3"><%=SplitCity(3)%><br>
                      ��<br>
                      ��</a></font></div>
                  </td>
                  <td height="50"> 
                    <p align="center"><font face="<%=SpecificFontFace%>" size="5"><a href="address.asp?c=4"><%=SplitCity(4)%><br>
                      ��<br>
                      ��</a></font></p>
                  </td>
                  <td height="50"> 
                    <div align="center"><font face="<%=SpecificFontFace%>" size="5"><a href="address.asp?c=5"><%=SplitCity(5)%><br>
                      ��<br>
                      ��</a></font></div>
                  </td>
                  <td height="50"> 
                    <div align="center"><font face="<%=SpecificFontFace%>" size="5"><a href="address.asp?c=6"><%=SplitCity(6)%><br>
                      ��<br>
                      ��</a></font></div>
                  </td>
                  <td height="50"> 
                    <div align="center"><font face="<%=SpecificFontFace%>" size="5"><a href="address.asp?c=0"><%=SplitCity(0)%><br>
                      ��<br>
                      ��</a></font></div>
                  </td>
                </tr>
              </table>
            </div>
	
<%
	End if
End if
%>
	   <br> 	
	  </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
'## �ر����ݿ�����
conn.close
set conn=nothing
%>
</body>
</html>

	
<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout.asp"
end if
%>

<%
Function Chked(id,value)
   '  To Check Check Boxes
   if id=value then
      Chked = "Selected"
   else 
      Chked = ""
   end if    
End Function
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
<script language="JavaScript">
<!--
function chkpwd() {
	pwd1=document.forms[0].passwd1.value;
	pwd2=document.forms[0].passwd2.value;
	if (pwd1!=pwd2) {
		alert("������������벻һ�£�");
		return;
	}
	if (pwd1=="") {
		alert("���벻��Ϊ�գ�");
		return;
	}
	document.forms[0].submit();
}
// -->
</script>
</head>
<body>
<%
profileid=request.QueryString("id")
If  profileid="" then
	if session("profile_id") <> "" then
		profileid = session("profile_id")	
	else
   		profileid=session("user_id")
   	end if
end if

if cint(profileid) <> cint(session("user_id")) and getUserLevel(session("user_id"))<3 then  
	Response.redirect("address.asp")
end if

session("profile_id")=profileid

mode=request.QueryString("mode")
If  mode="" then
	if session("profile_mode") <> "" then
		mode = session("profile_mode")	
	else
   		mode=1
   	end if
end if
session("profile_mode")=mode

If Request.ServerVariables("Request_Method")<>"GET" then
	mode=0
End if

set conn=server.createobject("adodb.connection")
conn.open ConnString
%>
<table width="760" align="center" border="0" height="300" cellspacing="2" cellpadding="0">
  <tr>
    <td width="118" valign="top" height="353"> 
      <table width="100%" border="0" height="353" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="118">
            <table width="100%" border="0">
              <tr bgcolor="#CCCCFF"> 
                <td height="25"><div align="center">
<% if mode=1 then%>
		�޸Ļ�������
<%else%>
                <a href="profile.asp?mode=1">�޸Ļ�������</a>
<%end if%>     
		</div></td>               
              </tr>
              <tr bgcolor="#CCCCFF"> 
                <td height="25"><div align="center">
<% if mode=2 then%>
		�޸ĸ��˼��
<%else%>
                <a href="profile.asp?mode=2">�޸ĸ��˼��</a>
<%end if%>                 
         </div></td>
              </tr>
              <tr bgcolor="#CCCCFF"> 
                <td height="25"><div align="center">
<% if mode=3 then%>
		�����û�ǩ��
<%else%>
                <a href="profile.asp?mode=3">�����û�ǩ��</a>
<%end if%>                 
                </div></td>                     
              </tr>
              <tr bgcolor="#CCCCFF"> 
                <td height="25"><div align="center">                
<% if mode=4 then%>
		�����û�����
<%else%>
                <a href="profile.asp?mode=4">�����û�����</a>
<%end if%>                 
                </div></td>                
              </tr>
              <tr bgcolor="#CCCCFF">
                <td height="25"><div align="center"><a href="address.asp">�� ��</a></div></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr valign="bottom"> 
          <td height="198"> 
            <table width="100%" border="0">
              <tr> 
                <td height="25" bgcolor="#FFCCCC"><font size="3">�û�����</font></td>
              </tr>
              <tr bgcolor="#CCFFCC"> 
                <td height="25"> 
                  <div align="center"><font color="#FF00FF" face="<%=SpecificFontFace%>"><% = getUserName(profileid) %></font></div>
                </td>
              </tr>
              <tr> 
                <td height="25" bgcolor="#FFCCCC"> 
                  <div align="left">�û�����</div>
                </td>
              </tr>
              <tr bgcolor="#CCFFCC"> 
                <td height="25"> 
                  <div align="center"><font color="#FF00FF" face="<%=SpecificFontFace%>"><% = getLevelName(getUserLevel(profileid)) %></font></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="640" align="center" height="353" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="3" style="border: 1px solid rgb(0,0,0)" height="360" bgcolor="<% =TableBgColor %>">
        <tr valign="top" align="center"> 
          <td><br>  
            
<%
if mode<>0 then
	strSQL  = " SELECT * FROM Members INNER JOIN UserInfo ON Members.Member_id = UserInfo.User_ID where Members.Member_id = " & profileid
	set rs = conn.Execute (strSQL)

	If rs.Eof or rs.Bof then  
		' No items found in DB
		Response.Write "û�з����������û���"
	Else
		txtsig=rs("U_sig")
		txtcomment=rs("U_Comment")
%>
	<form method="post" action="profile.asp">
<%            		
 		Select Case mode
			case 1 %>
        <table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr> 
            <td width="3%" height="35">&nbsp; </td>
            <td width="14%" height="35">�绰��</td>
            <td width="40%"> 
              <input type="text" name="phone" size="40" maxLength="100" value='<%=rs("U_Phone")%>'>*<a tabindex="-1" href="javascript:alert('˵����\r\r����绰����֮�����ö��Ÿ�����\r\r')"><font color='red' size="1">ע</font></a>*
            </td>
            <td width="20%" align="center"> 
                <% if getUserLevel(session("user_id"))=3 then %>
                   <table>
                     <tr>
                       <td>����</td>
                       <td><% =DoDropDown("UserLevel", "L_Name", "Level_ID", rs("M_Level"), "level", 0, 1) %></td>
                     </tr>
                   </table>
                <% end if%>
            </td>
          </tr>
          <tr> 
            <td width="3%" height="35">&nbsp;</td>
            <td width="14%" height="35">E-mail�� </td>
            <td colspan="2"> 
              <input type="text" name="email" size="40" maxLength="100" value='<% =rs("U_Email") %>'>*<a tabindex="-1" href="javascript:alert('˵����\r\r���Email֮�����÷ֺŸ�����\r\r')"><font color='red' size="1">ע</font></a>*
            </td>
          </tr>
          <tr> 
            <td width="3%" height="35">&nbsp;</td>
            <td width="14%" height="35">OICQ/ICQ��</td>
            <td colspan="2"> 
              <input type="text" name="oicq" size="30" maxLength="30" value='<% =rs("U_Oicq") %>'>
            </td>
          </tr>
          <tr> 
            <td width="3%" height="35">&nbsp;</td>
            <td width="14%" height="35">������ҳ��</td>
            <td colspan="2"> 
              <input type="text" name="homepage" size="40" maxLength="40" value='<% =rs("U_Homepage") %>'>
            </td>
          </tr>
          <tr> 
            <td width="3%" height="35">&nbsp;</td>
            <td width="14%" height="35">ͨѶ��ַ��</td>
            <td colspan="2"> 
              <input type="text" name="address" size="65" maxLength="100" value='<% =rs("U_Address") %>'>
            </td>
          </tr>
          <tr> 
            <td width="3%" height="35">&nbsp;</td>
            <td width="14%" height="35">�������룺</td>
            <td colspan="2"> 
              <input type="text" name="zipcode" size="8" maxLength="8" value='<% =rs("U_Zipcode") %>'>
            </td>
          </tr>
          <tr> 
            <td width="3%" height="35">&nbsp;</td>
            <td width="14%" height="35">���ڳ��У�</td>
            <td colspan="2">
            <% =DoDropDown("City", "C_Name", "City_ID", rs("U_city"), "city", 1, 1) %> 
            </td>
          </tr>
          <tr> 
            <td width="3%" height="24">&nbsp; </td>
            <td colspan="3" height="24"><font size="2">������ڳ���ѡ�����������������㵱ǰ���ڳ��У�</font> 
            <% if rs("U_city")=0 then %>    
              <input type="text" name="city2" maxLength="20" value='<% =rs("U_city2") %>'>    
            <%Else%>    
              <input type="text" name="city2" maxLength="20">    
            <%end if%>    
            </td>
          </tr>
          <tr> 
            <td width="3%" height="60">&nbsp; </td>
            <td colspan="3" height="60">
              <table width="100%" border="0">
                <tr> 
                  <td height="20" width="45%"> 
                    <div align="right"> 
                          <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:document.forms[0].submit()"><font size=2>ȷ ��</font></a></div>
                		</td>
              		      </tr>
            		   </table>                    
                    </div>
                  </td>
                  <td height="20" width="7%">&nbsp;</td>
                  <td height="20" width="48%"> 
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
            </td>
          </tr>
        </table>
<%			case 2 %>
        <table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr> 
            <td height="18" width="5%">&nbsp;</td>
            <td height="18" width="95%">���˼�飺*<a tabindex="-1" href="javascript:alert('                                      С���ɣ�����\r\r�������������м������·�������ǿ����Ч������������           \r�ԡ�^_^\r\r    [:)]��[:P]��[:(]��[;)]\r\r�������ϤHTML����Ҳ���������·�������ʾ������;����\r�Ŷ�Ӧ��ϵ���£�\r -------------------------------------------------------------------------\r      ����վ                        		       HTML\r -------------------------------------------------------------------------\r [b]��[/b]                             <b>��</b>\r [i]��[/i]                               <i>��</i>\r [a]��[/a]                            <a>��</a>\r [quote]��[/quote]             <BLOCKQUOTE>��</BLOCKQUOTE>\r [code]��[/code]                 <pre>��</pre>\r\r')"><font color='red' size="1">С����</font></a>*</td>
          </tr>
          <tr> 
            <td height="254" width="5%"> 
              <div align="center"> </div>
            </td>
            <td height="254" width="95%"> 
              <textarea name="comment" cols="76" rows="17" wrap="VIRTUAL"><%if txtcomment<>"" then response.write cleancode(txtcomment) %></textarea>
            </td>
          </tr>
          <tr> 
            <td height="43" colspan="2"> 
              <table width="100%" border="0">
                <tr> 
                  <td height="28" width="45%"> 
                    <div align="right"> 
                          <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:document.forms[0].submit()"><font size=2>ȷ ��</font></a></div>
                		</td>
              		      </tr>
            		   </table>
                    </div>
                  </td>
                  <td height="28" width="8%">&nbsp;</td>
                  <td height="28" width="47%"> 
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
            </td>
          </tr>
        </table>
  
<%			case 3 %>
        <table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr> 
            <td height="33" width="4%">&nbsp;</td>
            <td height="33" width="96%">�û�ǩ����*<a tabindex="-1" href="javascript:alert('                                      С���ɣ�����\r\r�������������м������·�������ǿ����Ч������������           \r�ԡ�^_^\r\r    [:)]��[:P]��[:(]��[;)]\r\r�������ϤHTML����Ҳ���������·�������ʾ������;����\r�Ŷ�Ӧ��ϵ���£�\r -------------------------------------------------------------------------\r      ����վ                        		       HTML\r -------------------------------------------------------------------------\r [b]��[/b]                             <b>��</b>\r [i]��[/i]                               <i>��</i>\r [a]��[/a]                            <a>��</a>\r [quote]��[/quote]             <BLOCKQUOTE>��</BLOCKQUOTE>\r [code]��[/code]                 <pre>��</pre>\r\r')"><font color='red' size="1">С����</font></a>*</td>
          </tr>
          <tr> 
            <td height="173" width="4%"> 
              <div align="center"></div>
            </td>
            <td height="173" width="96%"> 
              <textarea name="sig" cols="76" rows="15" wrap="VIRTUAL"><%if txtsig<>"" then response.write cleancode(txtsig) %></textarea>
            </td>
          </tr>
          <tr> 
            <td height="64" colspan="2"> 
              <table width="100%" border="0">
                <tr> 
                  <td height="2" width="45%"> 
                    <div align="right"> 
                          <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:document.forms[0].submit()"><font size=2>ȷ ��</font></a></div>
                		</td>
              		      </tr>
            		   </table>
                    </div>
                  </td>
                  <td height="2" width="8%">&nbsp;</td>
                  <td height="2" width="47%"> 
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
            </td>
          </tr>
        </table>
  
<%			case 4 %>

              <table width="54%" border="0" height="180">
                <tr> 
                  <td width="31%" height="20">&nbsp;</td>
                  <td width="69%">&nbsp; </td>
                </tr>
                <tr> 
                  <td width="31%">�����룺</td>
                  <td width="69%"> 
                    <input type="password" name="passwd1" size="20" maxLength="20" value='<% =rs("M_Password") %>'>
                  </td>
                </tr>
                <tr> 
                  <td width="31%">ȷ�����룺</td>
                  <td width="69%"> 
                    <input type="password" name="passwd2" size="20" maxLength="20" value='<% =rs("M_Password") %>'>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2"> 
                    <table width="100%" border="0">
                      <tr> 
                        <td height="34" width="45%"> 
                          <div align="right"> 
                          <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:chkpwd()"><font size=2>ȷ ��</font></a></div>
                		</td>
              		      </tr>
            		   </table>
                          </div>
                        </td>
                        <td height="34" width="7%">&nbsp;</td>
                        <td height="34" width="48%"> 
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
                  </td>
                </tr>
              </table>
<%		End Select %>
        <input name="m" type="hidden" value="<% =mode %>">
      </form>	
<%      
	End if
	
	rs.close
	set rs=nothing
Else
	If request.Form("m")=4 then
		strSql = "update Members set M_Password = '" & request.Form("passwd1")
		strSql = StrSql & "' where Member_id=" & profileid
		conn.Execute strSql
			
		if Err.description <> "" then 
			GO_Result "�޸Ĵ��� " & Err.description, false, "javascript:history.back()"
		Else
			Go_Result  "�����޸ĳɹ���", true, "address.asp"
		End If
	Else
		strSql= "U_Updated_By=" & session("user_id")
		strSql = strSql & ", U_Updated_Date=#" & now() & "#"
		
		Select Case request.Form("m")
			case 1
				if getUserLevel(session("user_id"))=3 then
					str = "update Members set M_Level = '" & request.Form("level")
					str = Str & "' where Member_id=" & profileid
					conn.Execute(str)
				end if
				
				strSql = strSql & ", U_Phone='" & trim(request.Form("phone")) 
				strSql = strSql & "', U_Email='" & trim(request.Form("email"))
				strSql = strSql & "', U_Oicq='" & trim(request.Form("oicq")) 
				strSql = strSql & "', U_Homepage='" & trim(request.Form("homepage")) 
				strSql = strSql & "', U_Address='" & trim(request.Form("address"))
				strSql = strSql & "', U_Zipcode='" & trim(request.Form("zipcode"))
				strSql = strSql & "', U_City=" & request.Form("city")
				strSql = strSql & ", U_City2='" & trim(request.Form("city2")) & "'"
				'strSql = strSql & ", U_Icon=" & request.Form("icon")
			case 2
				strSql = strSql & ", U_Comment='" & trim(chkString(request.Form("comment"))) & "'"
			case 3
				strSql = strSql & ", U_sig='" & trim(chkString(request.Form("sig"))) & "'"
		End Select
	
		strSql = "update UserInfo set " & StrSql & " where User_ID=" & profileid
		'Response.write strsql
		conn.Execute strSql
		
		if Err.description <> "" then 
			GO_Result "�޸Ĵ��� " & Err.description, false, "javascript:history.back()"
		Else
			Go_Result  "�����޸ĳɹ���", true, "address.asp"
		End If
	End if

End if
%>      	
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






<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout.asp"
end if
%>

<html>
<head>
<title>Ӱ��</title>
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

function chkphotofile() {
	if (document.forms[0].fname.value=="") {
		alert("�㻹û���ϴ���Ƭ�ļ���");
		return;
	}
	document.forms[0].submit();
}
// -->
</script>
</head>
<body><br>
<%
set conn=server.createobject("adodb.connection")
conn.open ConnString

If Request.ServerVariables("Request_Method")="GET" then

   method=lcase(request.QueryString("method"))

   If  method="" then
	method="new"
   End if

   strSql ="SELECT A_Name FROM Album where Album_ID = " & session("album_id") 
   'Response.write strSql
   set rs = conn.Execute (strSql)
   txtalbum=rs("A_Name")
   rs.close
   set rs=nothing
	
   If method = "edit" then
   	photoid=request.QueryString("photo_id")
	if  photoid="" then
		if session("photo_id") <> "" then
			photoid = session("photo_id")	
		else
   			Response.Redirect "album_info.asp"
   		end if
	end if
	session("photo_id")=photoid
	
	
	'##     �����ݿ�����
	strSql  = " SELECT * FROM Photoes where Photoes.Photo_ID = " &  photoid	
	set rs = conn.Execute (StrSql)
	
	txtfile=rs("P_Filename")
	txtsub = rs("P_Title")
	txtmsg=rs("P_Comment")
	txtdate =rs("P_Date")
	txtvert=rs("P_Vertical")
	
	rs.close
	set rs=nothing
   End if
   
'Response.write session("album_id") & ","
'Response.write session("photo_id") & ","
'Response.write session("user_id")
%>
<table width="760" border="0" cellspacing="2" cellpadding="0" align="center" style="border: 1px solid rgb(0,0,0)" bgcolor="<% =TableBgColor %>">
<tbody>
<tr>
<td>&nbsp;</td>
</tr>
<tr> 
  <th height="40" valign="top"><font face="<% =SpecificFontFace %>" color=#05006C size=4><% =txtalbum %></font></th>
</tr>
<tr>
<td align="center">
<form name="AttachForm" action="photo_post.asp" method="post">
   <table width="90%" border="0" cellspacing="0" cellpadding="3" align="center">
          <tr> 
            <td height="35"> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="78%">��Ƭ�ļ���
                    <input type="text" name="fnamealias" size="40" maxLength="40" value="<% =txtfile %>" disabled>
                  </td>
                  <td width="22%">
                  <%If method = "new" then%>	           
                    <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
                      <tr> 
                        <td> 
                          <div align="center" class="p9"><A href='#' onClick='WinOpen("upload.asp")'><font size=2>�ϴ��ļ�</font></a></div>
                        </td>
                      </tr>
                    </table>
                  <%End if%>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td height="29">��Ƭ���⣺     
              <input type="text" name="title" size="78" maxLength="78" value="<% =cleancode(txtsub) %>">
            </td>
          </tr>
          <tr valign="middle"> 
            <td height="32">�������ڣ�     
              <input type="text" name="date" maxLength="14" value="<% =txtdate %>">
            </td>
          </tr>
          <tr> 
            <td height="30">
              ��Ƭ˵����*<a tabindex="-1" href="javascript:alert('                                      С���ɣ�����\r\r�������������м������·�������ǿ����Ч������������           \r�ԡ�^_^\r\r    [:)]��[:P]��[:(]��[;)]\r\r�������ϤHTML����Ҳ���������·�������ʾ������;����\r�Ŷ�Ӧ��ϵ���£�\r -------------------------------------------------------------------------\r      ����վ                        		       HTML\r -------------------------------------------------------------------------\r [b]��[/b]                             <b>��</b>\r [i]��[/i]                               <i>��</i>\r [a]��[/a]                            <a>��</a>\r [quote]��[/quote]             <BLOCKQUOTE>��</BLOCKQUOTE>\r [code]��[/code]                 <pre>��</pre>\r\r')"><font color='red' size="1">С����</font></a>*</b></font> 
	    </td>
          </tr>
          <tr> 
            <td valign="top"> 
              <textarea name="comment" cols="90" rows="8" wrap="VIRTUAL"><%=cleancode(txtmsg)%></textarea>
            </td>
           </tr>
           <tr>
             <td>
		<font face="<% =DefaultFontFace %>" size="2"><input name="vertical" type="checkbox" value="true" <%=Chked(txtvert)%>>������ʾ��Ƭ</font>
            </td>
          </tr>
          <tr> 
            <td height="40" valign="bottom"> 
              <table width="100%" border="0">
                <tr> 
                  <td height="2" width="46%"> 
                    <div align="right"> 
                           <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
              		      <tr> 
                		<td> 
                  		  <div align="center" class="p9"><A href="javascript:chkphotofile()"><font size=2>ȷ ��</font></a></div>
                		</td>
              		      </tr>
            		   </table>
                    </div>
                  </td>
                  <td height="2" width="8%">&nbsp;</td>
                  <td height="2" width="46%"> 
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
   <INPUT type=hidden name=fname value="<% =txtfile %>">
   <input name="method" type="hidden" value="<% =method %>"> 
</form>	
</td>
</tr>
<tr>
<td>&nbsp;</td>
</tr>
</tbody></table>
<%Else
	title=trim(chkString(request.Form("title")))
	msg=trim(chkString(request.Form("comment")))
   	pdate=request.Form("date")
   	
	userid=Session("user_id")
	albumid=session("album_id")
	method=lcase(request.Form("method"))
	 
	if request.Form("vertical") <> "true" then
		TF  = "False"
	Else 
		TF = "True"
        End if
        		
	if method = "new" then
		strSql ="SELECT * from Photoes where P_Filename = '" & request.Form("fname") & "' and P_Album_ID = " & albumid
		'Response.Write StrSql
		set rs = conn.Execute (StrSql)

		If rs.Eof or rs.Bof then  
			strSql = "insert into Photoes (P_Filename, P_Title, P_Comment, P_Date, P_Posted_By, P_Album_ID, P_Vertical) Values ('"
			strSql = StrSql & request.Form("fname") & "', '"
			strSql = StrSql & title & "', '"
			strSql = StrSql & msg & "', '"
			strSql = StrSql & pdate & "', "
			strSql = StrSql & userid & ", "
			strSql = StrSql & albumid & ", "
			strSql = StrSql & TF & ")"
			'Response.write strSql
			conn.Execute strSql
		
			if Err.description <> "" then 
				GO_Result "���ݿ���� " & Err.description, false, "javascript:history.back()"
			Else
				Go_Result  "��Ƭ�ϴ��ɹ���", true, "album_info.asp"
			End If
		Else
			GO_Result "��ǰӰ�����Ѿ���ͬ��Ƭ�ļ���", false, "javascript:history.back()"
		End if
		
		rs.close
		set rs=nothing
	end if
	
	if method = "edit" then
		if session("photo_id") <> "" then
			photoid = session("photo_id")	
		else
			Response.Redirect "album_info.asp"
		end if
		
		strSql ="SELECT P_Posted_By from Photoes where Photo_ID = " & photoid 
		'Response.Write StrSql
		set rs = conn.Execute (StrSql)

		If rs.Eof or rs.Bof then  
			GO_Result "��Ƭ������!", false, "javascript:history.back()"
		Elseif chkPermission(albumid)=true then 
			msg = msg & vbcrlf & vbcrlf & " --- "& getUserName(userid) & " �޸���" & now()
			strSql = "update Photoes set P_Title = '" & title
			strSql = StrSql & "', P_Comment = '" & msg
			strSql = StrSql & "', P_Date = '" & pdate
			strSql = StrSql & "', P_Vertical = " & TF 
			strSql = strSql & " where Photo_ID=" & photoid
			'Response.write strsql
			conn.Execute strSql
			
			if Err.description <> "" then 
				GO_Result "���ݿ���� " & Err.description, false, "javascript:history.back()"
			Else
				Go_Result  "��Ƭ�޸ĳɹ���", true, "album_info.asp"
			End If
		Else	
			GO_Result "��û��Ȩ���޸�������Ƭ!", false, "javascript:history.back()"
		End if
		
		rs.close
		set rs=nothing
	end if
	
	'ˢ��Ӱ��ʱ��
	strSql = "update Album set A_Posted_Date=#" & now() & "# where Album_ID=" & albumid
	'Response.write strsql
	conn.Execute strSql
	
	'ˢ���û�ʱ��
	ownerid=getAlbumOwner(albumid)
	If ownerid<>0 then  
		strSql = "update UserInfo set U_Updated_Date=#" & now() & "# where User_ID=" & ownerid
		'Response.write strsql
		conn.Execute strSql		
	End if	
	
End if

'## �ر����ݿ�����
conn.close
set conn=nothing
%>
</body>
</html>
	

















<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout.asp"
end if
%>

<%
Function GetSig(id)
    strSQL = "Select U_Sig from UserInfo where User_ID = " & id 
    set rsSig = conn.Execute (strSQL)
    GetSig = rsSig("U_Sig")
    rsSig.close
    set rsSig = nothing
End Function

Sub DoEmail(email, user_name)
    	' ###  Emails Topic Author if Requested.  
    	' ###  This needs to be edited to use your own email component
    	' ###  if you don't have one, try the w3Jmail component from www.dimac.net it's free!
 
 	subject = PageTitle & "��̳���ӻظ�֪ͨ"
	msg = "��ã�" & user_name & "��" & vbcrlf & vbcrlf
	msg = msg & "����" & PageTitle & "��̳�����������˻ظ��ˣ���ȥ��������" & PageBaseHref & "���롣" & vbcrlf & vbcrlf
	msg = msg & "---------------------------" & vbcrlf
	msg = msg & "Hello! Your article in " & SiteName & " has just been answered. Press " & PageBaseHref & " to take a look." & vbcrlf
	
	Set JMail = Server.CreateObject("JMail.SMTPMail")
	JMail.AddHeader "Originating-IP","61.139.76.75" 
	'JMail.ServerAddress = "mail.abc.org"
	JMail.Sender = Webmaster
	JMail.AddRecipient email
	JMail.Subject = subject
	JMail.Body = msg
	JMail.Priority = 3
	JMail.Execute
	JMail.Close
	Set JMail = Nothing
End Sub
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

   if method = "edit" or method="reply" then
	if session("topic_id") <> "" then
		topicid = session("topic_id")	
	else
		Response.Redirect "forum.asp"
	end if
	
	'##     �����ݿ�����
	strSql = "SELECT Topics.T_subject, Topics.T_Message "
	strSql = strSql & "FROM Topics "
	strSql = strSql & "where Topics.Topic_id = " &  topicid

	set rs = conn.Execute (StrSql)
	
	txtsub = rs("T_subject")
	txtmsg = rs("T_Message")
	if method="reply" then
		txtsub = "RE: " & txtsub
		txtmsg =chr(10) & chr(13) & "  >> " & Replace(txtmsg, CHR(10), "  >> ") 
	end if
	
	rs.close
	set rs=nothing
   End if
%>
<form action="forum_post.asp" method="post">
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
<input maxLength="80" name="Subject" size="80" value="<% =cleancode(txtsub) %>">
</td></tr>
<tr>
<td noWrap vAlign="top"><font face="<% =DefaultFontFace %>" size="2"><b>����:<br>*<a tabindex="-1" href="javascript:alert('                                      С���ɣ�����\r\r�������������м������·�������ǿ����Ч������������           \r�ԡ�^_^\r\r    [:)]��[:P]��[:(]��[;)]\r\r�������ϤHTML����Ҳ���������·�������ʾ������;����\r�Ŷ�Ӧ��ϵ���£�\r -------------------------------------------------------------------------\r      ����վ                        		       HTML\r -------------------------------------------------------------------------\r [b]��[/b]                             <b>��</b>\r [i]��[/i]                               <i>��</i>\r [a]��[/a]                            <a>��</a>\r [quote]��[/quote]             <BLOCKQUOTE>��</BLOCKQUOTE>\r [code]��[/code]                 <pre>��</pre>\r\r')"><font color='red' size="1">С����</font></a>*</b></font> 
</td>
<td><textarea cols="80" name="Message" rows="10" wrap="VIRTUAL"><%=cleancode(txtMsg)%></textarea><br>
<font face="<% =DefaultFontFace %>" size="2">
<input name="Sig" type="checkbox" value="true" <%=Chked(session("sig"))%>>�����ҵ�ǩ����<br>
<%if method = "new" and getEmail(Session("user_id"))<>"" then%>
<input name="rmail" type="checkbox" value="true" <%=Chked(session("rmail"))%>>���˻ظ�ʱ�������ʼ�֪ͨ��
<%end if%>
</font></td></tr>
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
   txtsub=trim(chkString(request.Form("Subject")))
   txtMessage=trim(chkString(request.Form("message")))

   If  txtsub="" or txtMessage = "" then
	GO_Result "��������ݲ���Ϊ�գ�", false, "javascript:history.back()"
   else
	userid=Session("user_id")	
	method=lcase(request.Form("method"))
	session("sig")=request.Form("Sig")

	' #####
	if method = "new" then 
        	if Request.Form("sig") = "true" and GetSig(userid)<>"" then
        		txtMessage = txtMessage & vbcrlf & vbcrlf & "------------------"& vbcrlf & GetSig(userid)
        	End if
        	
        	session("rmail")=request.Form("rmail")
        	if Request.Form("rmail") <> "true" then
			TF  = "False"
		Else 
			TF = "True"
        	End if
	
		strSql = "insert into topics (T_Subject, T_Message, T_Originator, T_Mail) Values ('"
		strSql = StrSql & txtsub & "', '"
		strSql = StrSql & txtMessage & "', "
		strSql = StrSql & Userid & ", "
		strSql = StrSql & TF & ")"
		conn.Execute (StrSql)
		
		if Err.description <> "" then 
			GO_Result "�������ݴ��� " & Err.description, false, "javascript:history.back()"
		Else
			Go_Result  "���ӷ���ɹ���", true, "forum.asp?pageno=1"
		End IF
	End if

	
	if method = "edit" then
		if session("topic_id") <> "" then
			topicid = session("topic_id")	
		else
			Response.Redirect "forum.asp"
		end if
		
		strSql ="SELECT * from Topics where Topic_ID = " & topicid 
		'Response.Write StrSql
		set rs = conn.Execute (StrSql)
		
		If rs.Eof or rs.Bof then  
			GO_Result "���Ӳ�����!", false, "javascript:history.back()"
		Elseif rs("T_Originator")= userid or getUserLevel(userid)=3 then 
			'#  Do DB Update
			txtMessage = txtMessage & vbcrlf & vbcrlf & " --- "& getUserName(userid) & " �޸���" & now()
			
			if Request.Form("sig") = "true" and GetSig(userid)<>"" then
             			txtMessage = txtMessage & vbcrlf & vbcrlf & "------------------"& vbcrlf & GetSig(userid)
        		End if
        		
			strSql = "update Topics set T_subject = '" & txtsub
			strSql = StrSql & "', T_Message = '" & txtMessage
			strSql = StrSql & "' where topic_ID=" & topicid
			conn.Execute (StrSql)
	
			'# Update Last Post Time
			rootid=GetTopicID(topicid)
			strSql = "update Topics set T_Last_Posted = #" & now() & "# where Topic_ID = " & rootid
			conn.Execute (StrSql)
	
			if Err.description <> "" then 
				GO_Result "�޸Ĵ���" & Err.description, false, "javascript:history.back()"
			Else
				Go_Result  "�����޸ĳɹ���", true, "forum_info.asp"
			End If
		Else	
			GO_Result "��û��Ȩ���޸����ӣ�", false, "javascript:history.back()"
		End if
		
		rs.close
		set rs=nothing
	End if
	

	if method = "reply" then
		if session("topic_id") <> "" then
			topicid = session("topic_id")	
		else
			Response.Redirect "forum.asp"
		end if
		
		strSql ="SELECT * from Topics where Topic_ID = " & topicid 
		'Response.Write StrSql
		set rs = conn.Execute (StrSql)
		
		If rs.Eof or rs.Bof then  
			GO_Result "���Ӳ�����!", false, "javascript:history.back()"
		Else			
			if Request.Form("sig") = "true" and GetSig(userid)<>"" then
             			txtMessage = txtMessage & vbcrlf & vbcrlf & "------------------"& vbcrlf & GetSig(userid)
        		End if
	
			strSql = "insert into topics (T_Subject, T_Message, T_Originator, T_ParentID) Values ('"
			strSql = StrSql & txtsub & "', '"
			strSql = StrSql & txtMessage & "', "
			strSql = StrSql & Userid & ", "
			strSql = StrSql & topicid & ")"
			conn.Execute (StrSql)
		
			'# Update Last Post
			rootid=GetTopicID(topicid)
			strSql = "update Topics set T_Last_Posted = #" & now() & "#, T_Replies = T_Replies +1 where Topic_ID = " & rootid
			conn.Execute (StrSql)
	
			if Err.description <> "" then 
				GO_Result "�������ݴ��� " & Err.description, false, "javascript:history.back()"
			Else
				strSql ="SELECT Topics.T_Originator, Topics.T_Mail from Topics where Topic_ID = " & rootid 
				'Response.Write StrSql
				set rs1 = conn.Execute (strSql)
				if lcase(rs1("T_Mail")) = "true" then 
					strSQL  = " SELECT Members.M_Name, UserInfo.U_Email FROM Members INNER JOIN " & _ 
			   				" UserInfo ON Members.Member_id = UserInfo.User_ID WHERE Members.Member_id= " & rs1("T_Originator")
					set rs2 = conn.Execute (strSQL)
					if rs2("U_Email")<>"" then
						mail = split(rs2("U_Email"), ";")
						DoEmail  mail(0), rs2("M_Name")
					end if
					rs2.close
					set rs2 = nothing
				End if
				
				rs1.close
				set rs1 = nothing
				Go_Result  "���ӻظ��ɹ���", true, "forum_info.asp"
     			End if
     		End if
     		
     		rs.close
		set rs=nothing
     	End if
   End if
End if

'## �ر����ݿ�����
conn.close
set conn=nothing
%>
</body>
</html>
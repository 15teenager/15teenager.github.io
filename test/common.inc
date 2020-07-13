<!--#INCLUDE FILE="config.inc" -->

<%
On Error Resume Next

'set session timeout
Session.Timeout=20

'set buffering to true
response.buffer=true
		
Function getUserName(id)
if id = "" then
	getUserName = "无名氏"
else
	dim str2, my_conn, rs_user
	
	str2 ="SELECT M_Name FROM Members where Member_id = " & id
	
	set my_conn= Server.CreateObject("ADODB.Connection")
	my_conn.Open ConnString
	set rs_user = my_conn.Execute (str2)
	
	if rs_user.Eof or rs_user.Bof then
		getUserName = "无名氏"
	else
		getUserName = rs_user("M_Name")
	end if
	
	rs_user.close
	set rs_user=nothing
	my_conn.close
	set my_conn=nothing
end if
End Function 

Function getUserLevel(id)
if id = "" then
	getUserLevel = 0
else
	dim str2, my_conn, rs_user
	
	str2 ="SELECT M_Level FROM Members where Member_id = " & id
	
	set my_conn= Server.CreateObject("ADODB.Connection")
	my_conn.Open ConnString
	set rs_user = my_conn.Execute (str2)
	
	if rs_user.Eof or rs_user.Bof then
		getUserLevel = 0
	else
		getUserLevel = rs_user("M_Level")
	end if
	
	rs_user.close
	set rs_user=nothing
	my_conn.close
	set my_conn=nothing
end if	
End Function 

Function getLevelName(levelid)
	dim str2, my_conn, rs_level
	
	str2 ="SELECT L_Name FROM UserLevel where Level_ID = " & levelid
	
	set my_conn= Server.CreateObject("ADODB.Connection")
	my_conn.Open ConnString
	set rs_level = my_conn.Execute (str2)

	getLevelName = rs_level("L_Name")

	rs_level.close
	set rs_level=nothing
	my_conn.close
	set my_conn=nothing
End Function

Function getCity(cityid, city2)
	if cityid = 0 then 
		if city2<>"" then
			getCity = city2 
		else
			getCity = "<无>"
		end if
	else
		dim str2, my_conn, rs_city
		
		str2 ="SELECT C_Name FROM City where City_id = " & cityid
		
		set my_conn= Server.CreateObject("ADODB.Connection")
		my_conn.Open ConnString
		set rs_city = my_conn.Execute (str2)
		
		if rs_city.Eof or rs_city.Bof then
			getCity = "<无>"
		else
			getCity = rs_city("C_Name")
		end if
		
		rs_city.close
		set rs_city=nothing
		my_conn.close
		set my_conn=nothing
	end if
End Function

Sub GO_Result(msg, boolOk, url) 
%>
<br>
<center>
<font face="<% =DefaultFontFace %>" size="3"><p>
<%if boolOk = true then%>
<font color="<% =DefaultFontColor %>">
<%Else%>
<font color=red>
<%End If%>
<% =msg %></font></p></font>
<p>
<% if url<>"" then %>
      <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
        <tr> 
          <td> 
            <div align="center" class="p9"><a href="<%=url%>"><font size=2>返 回</font></a></div>
          </td>
        </tr>
      </table>
<% end if %>
</p>
</center>
<%
End Sub

'## 显示NEW图标的子程序
Function isNew(dt,page)
	Select Case page
		'case 1	
		'case 2
		'case 3
		'	newgif="<img src='gif/green-ball.gif' align='middle' alt='新' border=0>"
		'	oldgif="<img src='gif/yellow-ball.gif' align='middle' alt='' border=0>" 
		'case 4
		'	newgif="<img src='gif/new_t.gif' align='middle' alt='新' border=0>"
		'	oldgif="<img src='gif/old_t.gif' align='middle' alt='' border=0>" 
		case Else
			newgif="<img src='gif/icon_new.gif' align='middle' alt='新' border=0>"
			oldgif=""
	End Select
	
	if datediff("s", session("last_visited"), dt) > 1 then
		isNew = newgif
	Else
		isNew = oldgif
	End If
End Function

Function formatDate(dt)
	formatDate = year(dt)& "年"
	formatDate = formatDate & month(dt) & "月"
	formatDate = formatDate &  day(dt)  & "日 "
	if hour(dt)<10 then
		formatDate = formatDate & "0"
	end if
	formatDate = formatDate & hour(dt) & ":"
	if Minute(dt)<10 then
		formatDate = formatDate & "0"
	end if
	formatDate = formatDate & Minute(dt)
End Function

Function formatDate2(dt)
	formatDate2 = year(dt)& "年"
	formatDate2 = formatDate2 & month(dt) & "月"
	formatDate2 = formatDate2 &  day(dt)  & "日 "
End Function

Function FormatStr(String)
	on Error resume next
	String = Replace(String, CHR(13), "")
	String = Replace(String, CHR(10) & CHR(10), "</P><P>")
	String = Replace(String, CHR(10), "<BR>")
	FormatStr = String
End Function

Function CleanCode(str)
if str = "" then 
    str = ""
Else 
	str = replace(str, "<pre>", "[code]", 1, -1, 1)
	str = replace(str, "</pre>", "[/code]", 1, -1, 1)
	str = replace(str, "<b>", "[b]",1,-1,1)
	str = replace(str, "</b>", "[/b]",1,-1,1)
	str = replace(str, "<i>", "[i]",1,-1,1)
	str = replace(str, "</i>", "[/i]",1,-1,1)
	str = replace(str, "<BLOCKQUOTE><font size=1 face=arial>引述:<hr height=1 noshade>", "[quote]",1,-1,1)
	str = replace(str, "<hr height=1 noshade></BLOCKQUOTE></font><font face='" &DefaultFontFace& "' size=2>", "[/quote]",1,-1,1)
	str = replace(str, "<a href='", "[a]", 1, -1, 1)
	str = replace(str, "' Target=_Blank>链接</a>", "[/a]",1,-1,1)
	if smiles ="true" then
		str= replace(str, "<img src=""gif/wink.gif"" width=15 height=15 border=0 align=middle>", "[;)]",1,-1,1)
		str= replace(str, "<img src=""gif/sad.gif"" width=15 height=15 border=0 align=middle>", "[:(]",1,-1,1)
		str= replace(str, "<img src=""gif/tongue.gif"" width=15 height=15 border=0 align=middle>", "[:P]",1,-1,1)
		str= replace(str, "<img src=""gif/smile.gif"" width=15 height=15 border=0 align=middle>", "[:)]",1,-1,1)
		
	end if
	
	str = Replace(str, "&nbsp;", " ")
End if
CleanCode = str

end function

function doCode(str, oTag, cTag, roTag, rcTag)
	tx = split(str, cTag)
	t = ""
	for i = 0 to ubound(tx)
	  if lcase(oTag) =  "[a]" then
		p = instr(1, tx(i), "[a]", 1) 
		if p <> 0 then
			tmp = mid(tx(i), p)
			url = mid(tmp, 4)
			if lcase(left(url, 5)) = "http:" then
				tmp1 = Replace(tmp, "[a]"&url, "<A href='" & url & "' Target=_Blank>链接</a>", 1, -1, 1)			
			else
				tmp1 = Replace(tmp, "[a]"&url, "<A href='http://" & url & "' Target=_Blank>链接</a>", 1, -1, 1)	
			end if
			t =t & Replace(tx(i), tmp, tmp1)
		else
			t = t & tx(i)
		end if
	  else
		cnt = instr(1,tx(i), oTag,1)
		select case cnt 
			case 0
				t=t&tx(i)
			case else 
				t = t & Replace(tx(i), oTag, roTag,1,1,1)
				t = t & rcTag
		end select
	  end if
	next
	doCode = t
end function

Function Smile(string)
	 String = replace(String, "[:)]", "<img src=""gif/smile.gif"" width=15 height=15 border=0 align=middle>")
	 String = replace(String, "[:P]", "<img src=""gif/tongue.gif"" width=15 height=15 border=0 align=middle>")	
	 String = replace(String, "[:(]", "<img src=""gif/sad.gif"" width=15 height=15 border=0 align=middle>")	
	 String = replace(String, "[;)]", "<img src=""gif/wink.gif"" width=15 height=15 border=0 align=middle>")	
	 Smile = String
End function

Function ChkString(str)
	 if str = "" then 
		str = ""
	 Else
	 	str=server.htmlencode(str)
	 	
		if BadWordFiler = "true" then
		  bwords = split(BadWords, "|")
		  for i = 0 to ubound(bwords)
			str= replace(str, bwords(i), string(len(bwords(i)),"*"), 1,-1,1) 
		  next
        	End if
	 
	 	str = Replace(str, " ", "&nbsp;")
	 	
	 	'  Do ASP Forum Code
	 	str = doCode(str, "[code]", "[/code]", "<pre>", "</pre>")
		str = doCode(str, "[b]", "[/b]", "<b>", "</b>")
	 	str = doCode(str, "[i]", "[/i]", "<i>", "</i>") 
	 	str = doCode(str, "[quote]", "[/quote]", "<BLOCKQUOTE><font size=1 face=arial>引述:<hr height=1 noshade>", "<hr height=1 noshade></BLOCKQUOTE></font><font face='" & DefaultFontFace & "' size=2>") 
	 	str = doCode(str, "[a]", "[/a]", "<a>", "</a>")
	 	
	 	if smiles = "true" then str= smile(str)

	 	str = Replace(str, "'", "''")
	 	str = Replace(str, "|", "/")
	 
	 	ChkString = str
	 End if
End Function

Function GetTopicID(id)
	dim str2, my_conn, rs_topic
	
	str2 = "SELECT Topics.topic_id, Topics.T_ParentID "
	str2 = str2 & "FROM Topics "
	str2 = str2 & "where topics.topic_id = " &  id
	
	set my_conn= Server.CreateObject("ADODB.Connection")
	my_conn.Open ConnString
	set rs_topic = my_conn.Execute (str2)
	
	
	if rs_topic("T_ParentID")=0 then
		GetTopicID=rs_topic("topic_id")
	else
		GetTopicID=GetTopicID(rs_topic("T_ParentID"))
	end if
	
	rs_topic.close
	set rs_topic=nothing
	my_conn.close
	set my_conn=nothing
End Function

Function DoDropDown(tblName, DispField, ValueField, SelVal, name, t, numerical)
	dim StrSQL, my_conn, rs_drop
	
	StrSQL = "SELECT " &  DispField & ", " & ValueField 
	StrSQl = StrSQL & " FROM " & tblName 
	
	set my_conn= Server.CreateObject("ADODB.Connection")
	my_conn.Open ConnString
	'Response.Write strsql
	set rs_drop = my_conn.Execute (StrSQL)
	
	Response.Write "<Select Name='" & name & "'>"
	Select Case t
		case 0
			Response.Write ""  'Do nothing
		case 1
			Response.Write "<Option value='0'>其它</option>"  & vbcrlf
		case else 
			Response.Write "<Option value='0'>无</option>"  & vbcrlf
	End Select
	if numerical=1 then
		SelVal=cint(SelVal)
	end if
	do until rs_drop.EOF
		if rs_drop(ValueField) = SelVal then
			Response.Write "<option value='" & rs_drop(ValueField) & "' Selected>"
			Response.Write rs_drop(DispField) & "</option>" & vbcrlf
		Else
			Response.Write "<option value='" & rs_drop(ValueField) & "'>"
			Response.Write rs_drop(DispField) & "</option>" & vbcrlf
		End if
		rs_drop.MoveNext
	loop
	Response.Write "</select>" & vbcrlf
	
	rs_drop.close
	set rs_drop=nothing
	my_conn.close
	set my_conn=nothing
End Function

Function Chked(YN)
   '  To Check Check Boxes
   if YN or YN = "true" then
      Chked = "Checked"
   else 
      Chked = ""
   end if    
End Function

Function getAlbumID(id)
if id = "" then
	getAlbumID=0
else
	dim str, my_conn, rs2
	
	str ="SELECT * FROM Members INNER JOIN UserInfo ON Members.Member_id = UserInfo.User_ID where Members.Member_id =" & id
	
	set my_conn= Server.CreateObject("ADODB.Connection")
	my_conn.Open ConnString
	set rs2 = my_conn.Execute (str)
	
	If rs2.Eof or rs2.Bof then
		getAlbumID=0
	else
		getAlbumID=rs2("U_Album_ID")
	End if
	
	rs2.close
	set rs2=nothing
	my_conn.close
	set my_conn=nothing
end if
End Function

Function getAlbumOwner(id)
	dim str, my_conn, rs2
	
	str  = " SELECT User_ID FROM UserInfo where U_Album_ID = " & id
	'Response.write strsql
	
	set my_conn= Server.CreateObject("ADODB.Connection")
	my_conn.Open ConnString
	set rs2 = my_conn.Execute (str) 
	If rs2.Eof or rs2.Bof then
		getAlbumOwner=0
	else
		getAlbumOwner=rs2("User_ID")
	End if
	rs2.close
	set rs2=nothing
	my_conn.close
	set my_conn=nothing
End Function

Function chkPermission(id)
	ownerid=getAlbumOwner(id)
	If ownerid=0 then
		if getUserLevel(session("user_id"))>1 then
			chkPermission=true
		end if
	Else
		if ownerid = session("user_id") or getUserLevel(session("user_id"))=3 then
			chkPermission=true
		end if
	End if
End Function

Function getEmail(id)
	dim str2, my_conn, rs_mail
	
	str2 ="SELECT U_Email FROM UserInfo where User_ID = " & id
	
	set my_conn= Server.CreateObject("ADODB.Connection")
	my_conn.Open ConnString
	set rs_mail = my_conn.Execute (str2)

	getEmail = trim(rs_mail("U_Email"))

	rs_mail.close
	set rs_mail=nothing
	my_conn.close
	set my_conn=nothing
End Function
%>
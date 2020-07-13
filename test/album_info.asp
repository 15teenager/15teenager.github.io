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
<style>
body
	{background-image:URL("gif/album.gif")}
</style>
<script language="JavaScript">
<!--
function doConfirm(i)
{
  if (window.confirm("确实要删除吗？"))
    document.forms[i].submit();
}
// -->
</script>
</head>
<body>
<%
set conn=server.createobject("adodb.connection")
conn.open ConnString

albumid=request.QueryString("album_id")
if  albumid="" then
	if session("album_id") <> "" then
		albumid = session("album_id")	
	else
   		GO_Result "请指定影集！", false, ""
   	end if
end if
session("album_id")=albumid

strSql ="SELECT Album.A_Name, Album.A_Comment, Album.A_Directory FROM Album where Album.Album_ID = " & albumid
'Response.write strSql
set rs1 = conn.Execute (strSql)

If rs1.Eof or rs1.Bof then  
	' No items found in DB
	GO_Result "没有此影集！", false, ""
Else
	acom=rs1("A_Comment")
	permit=chkPermission(albumid)
%>

<table width="709" border="0">
  <tr>
    <td valign="top" align="left">
      <table align="left" border="0">
        <tr>
          <td width="73">&nbsp;</td>
          <td>影集标题：<font face='<%=SpecificFontFace%>' color=#05006C size=4><b><% =rs1("A_Name") %></b></font></td>
  	</tr>
  	<tr>
          <td width="73">&nbsp;</td>
          <td>影集说明：<br><font face='<%=SpecificFontFace%>' color=#05006C size=3><% if acom= "" then  Response.Write "<无>"  else  Response.Write formatStr(acom)  end if %></font><br>
          </td>
  	</tr>
  	<tr>
          <td width="73">&nbsp;</td>
          <td>&nbsp;</td>
  	</tr>
  	<tr>
          <td width="73">&nbsp;</td>
          <td><font size=2>影集目录：<% =rs1("A_Directory") %></font></td>
  	</tr>
      </table>
    </td>
    <td width="15%" valign="top">
    <form method="POST" action="album_erase.asp"> 
      <table width="100%" border="0">  
              <tr> 
                <td>        
            	  <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
                    <tr> 
                      <td> 
                        <div align="center" class="p9"><a href="javascript:window.close()"><font size=2>关闭</font></a></div>
                      </td>
              	    </tr>
            	  </table>	
                </td>
              </tr>            
              <% if permit=true  then %>
              <tr> 
                <td>        
            	  <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
                    <tr> 
                      <td> 
                        <div align="center" class="p9"><a href=album_post.asp?method=edit><font size=2>修改</font></a></div>
                      </td>
              	    </tr>
            	  </table>	
                </td>
              </tr>
              <tr> 
                <td> 
                  <%if getAlbumOwner(albumid)=0 then%>
                  <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
                    <tr> 
                      <td> 
                        <div align="center" class="p9"><a href="javascript:doConfirm(0)"><font size=2>删除</font></a></div>
                      </td>
              	    </tr>
            	  </table>
            	  <%end if%>
                </td>
              </tr>
              <%End if%>
      </table>
      </form>
    </td>
  </tr>
</table>
<br>
<hr align="center" width="650" noshade>
<% if permit=true  then %>
<table width="709" border="0">
  <tr>
    <td width="93">&nbsp;</td>
    <td>&nbsp;</td>
    <td width="28%" valign="top">       
      	  <table width="100%"><tr><td>
      	    <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
              <tr> 
                <td> 
                  <div align="center" class="p9"><a href=photo_post.asp><font size=2>上传照片</font></a></div>
                </td>
              </tr>
            </table>
      	  </td></tr></table>
    </td>
  </tr>
</table>
<%End if%>
<br>
<br>
<%
	'##     打开数据库连接
	strSql ="SELECT * FROM Photoes where Photoes.P_Album_ID = " & albumid & " order by Photoes.P_Posted_Date DESC"
	'Response.write strSql
	set rs = conn.Execute (strSql)
	
	If rs.Eof or rs.Bof then  
		' No items found in DB ""
		GO_Result "<img src='album/demo.gif' width='431' height='275' border='0'>", false, ""
	Else
		i=1
		rs.movefirst	
		do until rs.Eof
			pcom=rs("P_Comment")
			if rs("P_Vertical") then
				width=280
				height=431
			else
				width=431
				height=280
			end if
%>
<table width="709" border="0" height="275" cellspacing="0" cellpadding="0">
  <tr>
    <td width="93">&nbsp;</td>
    <td width="615">
      <table width="100%" border="0" cellspacing="2">
        <tr> 
          <td width="431" rowspan="5" align="center"> 
            <table width="<%=width%>" height="<%=height%>" border="0">
              <tr align="center"> 
                <td><a href='album/<%=rs1("A_Directory")%>/<%=rs("P_Filename")%>' target="_blank"><img src='album/<%=rs1("A_Directory")%>/<%=rs("P_Filename")%>' width="<%=width%>" height="<%=height%>" border="0"></a></td>
              </tr>
            </table>
          </td>
          <td width="28%" valign="top">标题：<font face='<%=SpecificFontFace%>' color=#05006C size=3><% if rs("P_Title")="" then Response.write "<无>" else Response.write rs("P_Title") end if%></font><% =isNew(rs("P_Posted_Date"),3) %></td>
        </tr>
        <% if rs("P_Date")<>"" then  Response.write "<tr><td width='28%' valign='top'><font size=2>&nbsp;&nbsp;摄于" & rs("P_Date") & "</font></td></tr>" end if%>
        <tr> 
          <td width="28%" valign="top">说明：<font face='<%=SpecificFontFace%>' color=#05006C size=3><% if pcom= "" then  Response.Write "<无>"  else  Response.Write formatStr(pcom)  end if %></font></td>
        </tr>
        <tr> 
          <td width="28%" height="10" valign="bottom"><font size=2>文件：<font color='#05006C'><% =rs("P_Filename") %></font></font><br><br>
        <font size=2>由<font color='<%=SpecificFontColor%>'><% =getUserName(rs("P_Posted_By")) %></font>上载于<% =formatDate2(rs("P_Posted_Date")) %></font></td>
        </tr>
        <tr>
          <td width="28%" valign="bottom"> 
            <% if permit=true then %>
            <form method="POST" action="photo_erase.asp"> 
            <INPUT type=hidden name=photo_id value="<%=rs("Photo_ID")%>">
            <table width="100%" border="0">              
              <tr> 
                <td> 
                  <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
              	    <tr> 
                      <td> 
                        <div align="center" class="p9"><a href=photo_post.asp?method=edit&photo_id=<%=rs("Photo_ID")%>><font size=2>修改</font></a></div>
                      </td>
              	    </tr>
            	  </table>	
                </td>
              </tr>
              <tr> 
                <td> 
                  <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
              	    <tr> 
                      <td> 
                        <div align="center" class="p9"><a href="javascript:doConfirm(<% =i %>)"><font size=2>删除</font></a></div>
                      </td>
              	    </tr>
            	  </table>
                </td>
              </tr>
            </table>
            </form>
            <%end if%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
	
<%
			rs.MoveNext
			i=i+1
  		loop
	  End if
	  rs.close
	  set rs=nothing
End if

rs1.close
set rs1=nothing
conn.close
set conn=nothing	  	
%>

</body>
</html>

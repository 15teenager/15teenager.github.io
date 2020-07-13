<!--#INCLUDE FILE="common.inc" -->

<%
if Session("user_id") = "" then
	Response.Redirect "timeout2.asp"
end if
%>

<html>
<head>
<title>上传照片</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
<script language="JavaScript">
<!--
function Done(){
	fname=TN(document.UploadFileForm.fdata.value);
	if (fname=="") 
	{
		alert("照片文件名不能为空！");
		return;
	}
	if(validateFileEntry(fname)==true) {
		document.UploadFileForm.fn.value = fname;
		window.opener.document.forms[0].fnamealias.value = fname;
		window.opener.document.forms[0].fname.value = fname;
		document.UploadFileForm.submit();
	}
}

function TN(Src){
	var Index = Src.lastIndexOf("\\");
	if (Index==-1) 
		Index = Src.lastIndexOf("\/");
	return Src.substring(Index+1,Src.length);
}

function validateFileEntry(validString) {
	var isCharValid = true;
	var inValidChar;
	for (i=0 ; i < validString.length ; i++) {
		if (validString.charAt(0) == '.') {
			isCharValid = false;
			i=validString.length;
		}
		if (validateCharacter(validString.charAt(i)) == false) {
			isCharValid = false;
			inValidChar = validString.charAt(i);
			i=validString.length;
		}
	}	   
	if (i < 1) { return false; }	   
	if (isCharValid == false) {
		if (inValidChar) { alert("无效的文件名. 不能含有 '" + inValidChar + "'.");	}
		else             { alert("无效的文件名. 请重新输入."); }
		return false;
   }
   return true;
}
	
function validateCharacter(character) {
   if ((character >= 'a' && character <= 'z') || ( character >='A' && character <='Z') || ( character >= '0' && character <= '9') || ( character =='-') || ( character == '.') || ( character == '_')) return true;	
   else return false;
}
// -->
</script>
</head>
<body>
<%
set conn=server.createobject("adodb.connection")
conn.open ConnString

strSql ="SELECT A_Directory FROM Album where Album_ID = " & session("album_id") 
'Response.write strSql
set rs = conn.Execute (strSql)
%>
<table width="450" height="310" align="center" border="0" cellspacing="2" cellpadding="5"> 
  <tr> 
    <td>     
      <form name="UploadFileForm" ENCTYPE="multipart/form-data" method="post" action="upload.cgi" >
        <INPUT type=hidden name=albumdir value='<% =rs("A_Directory") %>'>
        <INPUT type=hidden name=fn>
        <table border="0" width="100%" align="center">
          <tr> 
      	    <th colspan="2"><font face="<% =SpecificFontFace %>">上传照片</font></th>
    	  </tr>
    	  <tr> 
      	    <td colspan="2">&nbsp;</td>
    	  </tr>
          <tr> 
            <td width="19%">照片文件：</td>
            <td width="81%"> 
              <input type="file" name="fdata" size="35">
            </td>
          </tr>
          <tr> 
            <td colspan="2"> 
              <table width="100%" border="0">
                <tr> 
                  <td width="45%" align="right" valign="middle" height="58"> 
                    <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
                      <tr> 
                        <td> 
                          <div align="center" class="p9"><A href="javascript:Done()"><font size=2>确 认</font></a></div>
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td width="15%" height="58">&nbsp;</td>
                  <td width="45%" align="left" valign="middle" height="58"> 
                    <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="left">
                      <tr> 
                        <td> 
                          <div align="center" class="p9"><A href="javascript:window.close()"><font size=2>关 闭</font></a></div>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr valign="bottom"> 
            <td colspan="2" height="2"> 
              <hr width="100%" size="1" align="left">
            </td>
          </tr>
          <tr valign="top"> 
            <td colspan="2"> 
              <p>照片文件要求：<br>
                &nbsp;&nbsp;&nbsp;&nbsp;1。请采用*.jpg文件格式或*.gif文件格式；<br>
                &nbsp;&nbsp;&nbsp;&nbsp;2。建议文件大小不超过60Kb；<br>
                &nbsp;&nbsp;&nbsp;&nbsp;3。照片尺寸最小430*280像素，最大720*450像素。 </p>
              <p><font color="#FF0000">注意：上传一个照片文件可能需要几秒到几分钟的时间，由网络连线的速度和照片文件的大小决定，请耐心等待 
                ......<br>
                </font></p>
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
<%
rs.close
set rs=nothing
		
'## 关闭数据库连接
conn.close
set conn=nothing
%>
</body>
</html>

	
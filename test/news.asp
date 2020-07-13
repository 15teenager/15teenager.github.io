<!--#INCLUDE FILE="common.inc" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
<script language="JavaScript">
<!--
function MM_timelinePlay(tmLnName, myID) { //v1.2
  //Copyright 1997 Macromedia, Inc. All rights reserved.
  var i,j,tmLn,props,keyFrm,sprite,numKeyFr,firstKeyFr,propNum,theObj,firstTime=false;
  if (document.MM_Time == null) MM_initTimelines(); //if *very* 1st time
  tmLn = document.MM_Time[tmLnName];
  if (myID == null) { myID = ++tmLn.ID; firstTime=true;}//if new call, incr ID
  if (myID == tmLn.ID) { //if Im newest
    setTimeout('MM_timelinePlay("'+tmLnName+'",'+myID+')',tmLn.delay);
    fNew = ++tmLn.curFrame;
    for (i=0; i<tmLn.length; i++) {
      sprite = tmLn[i];
      if (sprite.charAt(0) == 's') {
        if (sprite.obj) {
          numKeyFr = sprite.keyFrames.length; firstKeyFr = sprite.keyFrames[0];
          if (fNew >= firstKeyFr && fNew <= sprite.keyFrames[numKeyFr-1]) {//in range
            keyFrm=1;
            for (j=0; j<sprite.values.length; j++) {
              props = sprite.values[j]; 
              if (numKeyFr != props.length) {
                if (props.prop2 == null) sprite.obj[props.prop] = props[fNew-firstKeyFr];
                else        sprite.obj[props.prop2][props.prop] = props[fNew-firstKeyFr];
              } else {
                while (keyFrm<numKeyFr && fNew>=sprite.keyFrames[keyFrm]) keyFrm++;
                if (firstTime || fNew==sprite.keyFrames[keyFrm-1]) {
                  if (props.prop2 == null) sprite.obj[props.prop] = props[keyFrm-1];
                  else        sprite.obj[props.prop2][props.prop] = props[keyFrm-1];
        } } } } }
      } else if (sprite.charAt(0)=='b' && fNew == sprite.frame) eval(sprite.value);
      if (fNew > tmLn.lastFrame) tmLn.ID = 0;
  } }
}

function MM_timelineGoto(tmLnName, fNew, numGotos) { //v2.0
  //Copyright 1997 Macromedia, Inc. All rights reserved.
  var i,j,tmLn,props,keyFrm,sprite,numKeyFr,firstKeyFr,lastKeyFr,propNum,theObj;
  if (document.MM_Time == null) MM_initTimelines(); //if *very* 1st time
  tmLn = document.MM_Time[tmLnName];
  if (numGotos != null)
    if (tmLn.gotoCount == null) tmLn.gotoCount = 1;
    else if (tmLn.gotoCount++ >= numGotos) {tmLn.gotoCount=0; return}
  jmpFwd = (fNew > tmLn.curFrame);
  for (i = 0; i < tmLn.length; i++) {
    sprite = (jmpFwd)? tmLn[i] : tmLn[(tmLn.length-1)-i]; //count bkwds if jumping back
    if (sprite.charAt(0) == "s") {
      numKeyFr = sprite.keyFrames.length;
      firstKeyFr = sprite.keyFrames[0];
      lastKeyFr = sprite.keyFrames[numKeyFr - 1];
      if ((jmpFwd && fNew<firstKeyFr) || (!jmpFwd && lastKeyFr<fNew)) continue; //skip if untouchd
      for (keyFrm=1; keyFrm<numKeyFr && fNew>=sprite.keyFrames[keyFrm]; keyFrm++);
      for (j=0; j<sprite.values.length; j++) {
        props = sprite.values[j];
        if (numKeyFr == props.length) propNum = keyFrm-1 //keyframes only
        else propNum = Math.min(Math.max(0,fNew-firstKeyFr),props.length-1); //or keep in legal range
        if (sprite.obj != null) {
          if (props.prop2 == null) sprite.obj[props.prop] = props[propNum];
          else        sprite.obj[props.prop2][props.prop] = props[propNum];
      } }
    } else if (sprite.charAt(0)=='b' && fNew == sprite.frame) eval(sprite.value);
  }
  tmLn.curFrame = fNew;
  if (tmLn.ID == 0) eval('MM_timelinePlay(tmLnName)');
}

function MM_initTimelines() {
    //MM_initTimelines() Copyright 1997 Macromedia, Inc. All rights reserved.
    var ns = navigator.appName == "Netscape";
    document.MM_Time = new Array(1);
    document.MM_Time[0] = new Array(2);
    document.MM_Time["Timeline1"] = document.MM_Time[0];
    document.MM_Time[0].MM_Name = "Timeline1";
    document.MM_Time[0].fps = 6;
    document.MM_Time[0][0] = new String("sprite");
    document.MM_Time[0][0].slot = 1;
    if (ns)
        document.MM_Time[0][0].obj = document["Layer1"];
    else
        document.MM_Time[0][0].obj = document.all ? document.all["Layer1"] : null;
    document.MM_Time[0][0].keyFrames = new Array(1, 12, 25, 44, 64, 75, 87, 96, 109, 121);
    document.MM_Time[0][0].values = new Array(2);
    document.MM_Time[0][0].values[0] = new Array(544,549,554,559,565,570,575,579,583,586,588,589,589,589,588,588,586,585,583,580,576,571,564,555,554,554,553,553,552,560,559,557,555,543,531,519,508,496,485,474,464,453,443,434,428,423,419,417,415,414,413,413,413,413,412,412,412,411,410,407,402,395,386,375,349,323,295,267,238,206,174,142,109,75,40,27,25,27,31,37,43,50,57,65,73,82,92,110,119,127,133,140,147,156,168,185,201,215,229,242,255,268,281,295,310,325,341,358,376,397,418,440,462,485,507,530,552,545,538,520,543);
    document.MM_Time[0][0].values[0].prop = "left";
    document.MM_Time[0][0].values[1] = new Array(19,38,57,77,96,115,134,153,171,190,207,225,238,251,263,275,287,298,309,320,331,332,332,340,348,342,344,345,346,346,346,346,345,343,342,340,348,345,343,340,346,343,339,334,339,334,327,320,313,306,299,292,285,278,272,265,258,250,242,234,226,220,217,216,222,234,248,264,281,299,317,335,341,343,350,334,302,270,238,207,177,147,119,93,68,45,23,27,47,64,81,96,112,127,141,148,144,136,126,115,103,91,78,66,54,42,31,22,15,10,7,6,5,5,6,7,8,10,11,13,14);
    document.MM_Time[0][0].values[1].prop = "top";
    if (!ns) {
        document.MM_Time[0][0].values[0].prop2 = "style";
        document.MM_Time[0][0].values[1].prop2 = "style";
    }
    document.MM_Time[0][1] = new String("behavior");
    document.MM_Time[0][1].frame = 122;
    document.MM_Time[0][1].value = "MM_timelineGoto('Timeline1','1')";
    document.MM_Time[0].lastFrame = 122;
    for (i=0; i<document.MM_Time.length; i++) {
        document.MM_Time[i].ID = null;
        document.MM_Time[i].curFrame = 0;
        document.MM_Time[i].delay = 1000/document.MM_Time[i].fps;
    }
}
//-->
</script>
</head>
<BODY onLoad="MM_timelinePlay('Timeline1')">
<%
set conn=server.createobject("adodb.connection")
conn.open ConnString

'记录用户动作
strSql = "insert into Events (L_Visited_By, L_Event, L_ScriptName, L_QueryString, L_Referer) Values ("
strSql = StrSql & session("user_id") & ", '"
strSql = StrSql & "浏览" & "', '"
strSql = StrSql & request.servervariables("script_name") & "', '"
strSql = StrSql & request.QueryString & "', '"
strSql = StrSql & Request.ServerVariables("HTTP_REFERER") & "')"
'Response.write strsql
'conn.Execute strSql

mypage=request.QueryString("pageno")
If  mypage="" then
	if session("news_pageno") <> "" then
		mypage = session("news_pageno")	
	else
   		mypage=1
   	end if
end if
session("news_pageno")=mypage

'选择最新用户设为动画显示
strSql = "select * from Album INNER JOIN UserInfo ON Album.Album_ID = UserInfo.U_Album_ID order by UserInfo.U_Updated_Date Desc"
set rs = conn.Execute (StrSql)

strSql ="SELECT * FROM Photoes where P_Album_ID = " & rs("Album_ID") & " order by P_Posted_Date Desc"
set rs1 = conn.Execute (StrSql)

if rs1.bof or rs1.eof then
	picpath = "album/demo.gif"
else
	picpath = "album/" & rs("A_Directory") & "/" & rs1("P_Filename")
end if

rs1.close
set rs1=nothing
%>
<div id="Layer1" style="position:absolute; width:83px; height:83px; z-index:1; left: 644px; top: 19px"> 
  <table width="100%" border="0" height="71" cellspacing="0" cellpadding="0">
    <tr bgcolor="#FF00FF"> 
      <td height="32"> 
        <div align="center">
	  <b><font color="#FFFFFF" size="3" face="楷体_GB2312">每周一<img src="gif/star.gif" width="30" height="29" align="absmiddle"></font></b>
	  <br>
	  <b><font color="#FFFF00" size="3" face="楷体_GB2312">☆<%=getUserName(rs("User_ID"))%>☆</font></b>	
	</div>
      </td>
    </tr>
    <tr> 
      <td> 
        <div align="center"><a href='address.asp?id=<%=rs("User_ID")%>' target='fbody'><img src='<%=picpath%>' width="100" height="75" border="0"></a></div>
      </td>
    </tr>
  </table>
</div>

<table width="645" border="0" cellspacing="0" cellpadding="3" style="border: 1px solid rgb(0,0,0)" height="360" bgcolor="<% =TableBgColor %>"> 
  <tr>
    <td valign="top"> 
      <% if getUserLevel(session("user_id"))>1 then %>
      	  <table width="100%"><tr><td align=right>
      	    <table border="1" cellspacing="0" cellpadding="2" bgcolor="<% =ButtonBgColor %>" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="right">
              <tr> 
                <td> 
                  <div align="center" class="p9"><a href=news_post.asp target='fbody'><font size=2>发布新闻</font></a></div>
                </td>
              </tr>
            </table>
      	  </td></tr></table>
      <%End if%>
      <br>    
      <table width="100%" border="0" cellspacing="0" cellpadding="3">

<%
'##     打开数据库连接
strSql ="SELECT News.News_ID, News.N_Title, News.N_Posted_Date "
strSql = strSql & "FROM News "
strSql = strSql & "order by News.N_Posted_Date DESC"

set rs=server.createobject("adodb.recordset")
rs.cachesize=Pagesize
rs.open strSql,conn,3

'显示所有Topic
If rs.Eof or rs.Bof then  
	' No items found in DB
	Response.Write "<tr><td width='20'>&nbsp;</td><td>没有新闻！</td></tr>"
Else
	rs.movefirst
	rs.pagesize=Pagesize
	maxpages=cint(rs.pagecount)
	if cint(mypage) > maxpages then 
		mypage=maxpages
		session("news_pageno")=mypage
	end if
	rs.absolutepage=mypage

	rec = 1
	do until rs.Eof or rec > Pagesize 		
		'## Display News
	  	Response.Write "<tr>"
	  	Response.Write "<td width='20' align='center'>" & isNew(rs("N_Posted_Date"),1) & "</td>" & vbcrlf
	  	Response.Write "<td><font face='" & DefaultFontFace & "' color= '" & DefaultFontColor & "' size='2'><a href='news_info.asp?news_id=" & rs("news_ID") & "' target='fbody'><img src='gif/green-ball.gif' border=0>" & rs("N_Title") & "</a>&nbsp;........&nbsp;"  & rs("N_Posted_Date") & "</font></td>"
	  	Response.Write "</tr>"
	  	rs.MoveNext
	    	
	    	rec = rec + 1
	loop
End If

rs.close
set rs=nothing
%>
      </table>
      
<%if maxpages > 1 then  %>
<br><hr width="95%" size="2" align="center">
<table width="93%" align="center">
<tr><td>
<div align=left><font face="<% =DefaultFontFace %>" size="2">
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
</td></tr></table>
<%End if%>      
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



<%
'****************************************
'*      ���²�����Ҫ�޸�                *
'****************************************

PageTitle = "ABC�༶"
PageBaseHref="http://www.abc.org/"
SiteName = "abc"
Webmaster="zhijun_wong@yahoo.com"



'****************************************
'*      ���²������Բ����޸�            *
'****************************************

ConnString = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & server.mappath("db/alumni.mdb")
Pagesize=30
Linesize=25

PageBGColor = "#CCFFFF"
TableBgColor = "#EEEEFF"
ButtonBgColor = "#CCCCCC"
DefaultBgColor="#CCCCCC"

LogoImgLocation="gif/logo.gif"
BannerImgLocation = "gif/banner.jpg"
DigitPath = "digits/red_on_white"

DefaultFontFace = "Verdana, Arial, Helvetica"
DefaultFontColor = "#000080"
SpecificFontFace = "����_GB2312"
SpecificFontColor = "#FF00FF"

HeadCellColor="#0000A0"
HeadFontColor = "#ffffff"

ForumCellColor = "#f7f7f7"
AltForumCellColor = "#e1e1e1"

ForumFontColor = "#000080"
ForumLinkColor = "#000080"

NewFontColor = "#ff0000"

Smiles = "true"  '  True/False to use Smiles

BadWordFiler = "true"   '  True / False To use filter
BadWords = "fuck|wank|shit|ȥ����|���˵�|"
%>
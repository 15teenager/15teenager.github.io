#!d:/perl/bin/perl

####################################
$rootdir=&mypath;

####################################
use vars qw($in);
use CGI  qw(:cgi);
$in = new CGI;

my $donepage = $in->param('location');
my $pathname = $in->param('path');
my $filename = $in->param('file');
my $fullpath = "$rootdir/album/$pathname";

if ($filename) {
	if (-e "$fullpath/$filename") {
		unlink ("$fullpath/$filename");
	}
	if ($donepage) {
		print "Location: $donepage\n\n";
	}
	else {
		&output ("/album_info.asp", "返 回");
	}
}
else {
	if (-e $fullpath){
		opendir (DIR, $fullpath) or exit;
		my @ls = readdir(DIR);
		closedir (DIR);

		my $file;
		FILE: foreach $file (@ls) {
			# Skip the "." entry and ".." if we are at root level.
			next FILE if  ($file eq '.');
			next FILE if  ($file eq '..');
		
			unlink ("$fullpath/$file");
		}
		rmdir($fullpath);
	}
	if ($donepage) {
		print "Location: $donepage\n\n";
	}
	else {
		&output ("javascript:window.close()", "关 闭");
	}
}


sub mypath
{
	$tempfilename=__FILE__;
	if ($tempfilename=~/\\/) { $tempfilename=~ s/\\/\//g;}

	if ($tempfilename) {
	   $mypath=substr($tempfilename,0,rindex($tempfilename,"/"));
	  }
	else {
	   $mypath=substr($ENV{'PATH_TRANSLATED'},0,rindex($ENV{'PATH_TRANSLATED'},"\\"));
	   $mypath=~ s/\\/\//g;
	  }
}


sub output{
local ($link, $btn) = @_;

$|++;						# Flush Output
print "content-type:text/html\n\n";

$envtext_tab.= <<EOFXX;
<html>
<head>
<title>影集</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
</head>
<body><br>
<table width="450" align="center" border="0" cellspacing="2" cellpadding="5"> 
  <tr>
    <td align=center>&nbsp;</td>
  </tr>
  <tr>
    <td align=center><font>删除成功！</font></td>
  </tr>
  <tr> 
    <td align=center height="58">                
                    <table border="1" cellspacing="0" cellpadding="2" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
                      <tr> 
                        <td> 
                          <div align="center" class="p9"><A href=$link><font size=2>$btn</font></a></div>
                        </td>
                      </tr>
                    </table>
    </td>
  </tr>  
</table> 
</body> 
</html>         
EOFXX
$envtext=$envtext_head.$envtext_tab.$envtext_var.$envtext2;
print $envtext;
}

#!d:/perl/bin/perl
#####################################
### 网友自荐Beta1.0
### 程式制作: CGI工作室
### 网址: bbss.126.com
### Email: bbss@kali.com.cn
### 请访问本站获得更多的CGI程序
### 禁止商业使用，商业定义--一切和有收入相关的网站(包括放置第二方盈利性广告）
### 如有问题或BUG，请来信指正，本站将根据问题加以解决。
####################################
## 设置开始##
## Note: NO trailing slash
$rootdir=&mypath;

## Do you wish to allow all file types?  yes/no (no capital letters)
$allowall = "no";      # 是否上传所有的文件类型

## If the above = "no"; then which is the only extention to allow?
## Remember to have the LAST 4 characters i.e. .ext
$theext = ".jpg";       #如果上面设置不，可以上传的文件类型
$thee = ".gif";

## The page you wish it to forward to when done:
## I.E. http://www.mydomainname.com/thankyou.html
#$donepage = "http://ads/";          #上传成功后自动转向地址

################################################
use vars qw($in);
use CGI  qw(:cgi);
$in = new CGI;

$|++;	# Flush Output
print "content-type:text/html\n\n";

my $fileName= $in->param("fn"); 
if ($fileName eq "") {
	&error("文件名不能为空！");
	exit 0;
}
else {
	#检查文件类型
	if ($allowall eq "yes") {
		$filetypeok = "yes";
	}
	else {
		my $ext=lc(substr($fileName,length($fileName) - 4,4));
		if ($ext eq $theext){
			$filetypeok = "yes";
		}
		if ($ext eq $thee){
			$filetypeok = "yes";
		}
	}
	
	if ($filetypeok eq "yes") {
		my $imagedir=$in->param("albumdir");
		$imagedir="$rootdir/album/$imagedir";
		
		#如果目录不存在，就创建它。
		if (!&exists($imagedir)) {
			if(!mkdir ($imagedir, 0755)) {
				&error("建立目录失败，请重试！");
				exit 0;
			}
		}
		
		my $fullfile= "$imagedir/$fileName";
		#print $fullfile;
		if (&exists($fullfile)) { #检查同名文件是否存在
			&error("“$fileName”已经存在,请改变文件名，然后再重新上传！");	
			exit 0;
		}
		else { #创建上传文件
			my $file = $in->param("fdata"); 
			my $file_size = 0;
			
			open (OUTFILE, ">$fullfile"); 
			binmode (OUTFILE);	# For those O/S that care.
			while (my $bytesread = read($file, my $buffer, 1024)) { 
				print OUTFILE $buffer;
				
				$file_size += $bytesread;
				if (($file_size / 1000) > 60) {
					close OUTFILE;
					unlink ($fullfile);
					&error("上传文件失败！你指定的文件大于我们限定的文件大小60Kb!");
					exit 0;
				}
			} 
			close (OUTFILE);
			#print "文件已上传";
			print "<script language='javascript'>window.close();</script>";
		}
	}
	else {
		&error("文件格式不正确,请上传*.jpg文件或*.gif文件。");
		exit 0;
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


sub exists {
# -----------------------------------------------------
# Checks to see if a file exists.
#	
	return -e shift;
}


sub error {
local ($errors) = @_;
$envtext_tab.= <<EOFXX;
<html>
<head>
<title>错误</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type=text/css href='style.css'>
</head>
<body><br>
<table width="450" align="center" border="0" cellspacing="2" cellpadding="5"> 
  <tr>
    <td align=center>&nbsp;</td>
  </tr>
  <tr>
    <td align=center><font color="red">$errors</font></td>
  </tr>
  <tr> 
    <td align="center" height="58">
                    <table border="1" cellspacing="0" cellpadding="2" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="80" align="center">
                      <tr> 
                        <td> 
                          <div align="center" class="p9"><A href="javascript:history.back()"><font size=2>返 回</font></a></div>
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
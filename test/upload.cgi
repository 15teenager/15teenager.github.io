#!d:/perl/bin/perl
#####################################
### �����Լ�Beta1.0
### ��ʽ����: CGI������
### ��ַ: bbss.126.com
### Email: bbss@kali.com.cn
### ����ʱ�վ��ø����CGI����
### ��ֹ��ҵʹ�ã���ҵ����--һ�к���������ص���վ(�������õڶ���ӯ���Թ�棩
### ���������BUG��������ָ������վ������������Խ����
####################################
## ���ÿ�ʼ##
## Note: NO trailing slash
$rootdir=&mypath;

## Do you wish to allow all file types?  yes/no (no capital letters)
$allowall = "no";      # �Ƿ��ϴ����е��ļ�����

## If the above = "no"; then which is the only extention to allow?
## Remember to have the LAST 4 characters i.e. .ext
$theext = ".jpg";       #����������ò��������ϴ����ļ�����
$thee = ".gif";

## The page you wish it to forward to when done:
## I.E. http://www.mydomainname.com/thankyou.html
#$donepage = "http://ads/";          #�ϴ��ɹ����Զ�ת���ַ

################################################
use vars qw($in);
use CGI  qw(:cgi);
$in = new CGI;

$|++;	# Flush Output
print "content-type:text/html\n\n";

my $fileName= $in->param("fn"); 
if ($fileName eq "") {
	&error("�ļ�������Ϊ�գ�");
	exit 0;
}
else {
	#����ļ�����
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
		
		#���Ŀ¼�����ڣ��ʹ�������
		if (!&exists($imagedir)) {
			if(!mkdir ($imagedir, 0755)) {
				&error("����Ŀ¼ʧ�ܣ������ԣ�");
				exit 0;
			}
		}
		
		my $fullfile= "$imagedir/$fileName";
		#print $fullfile;
		if (&exists($fullfile)) { #���ͬ���ļ��Ƿ����
			&error("��$fileName���Ѿ�����,��ı��ļ�����Ȼ���������ϴ���");	
			exit 0;
		}
		else { #�����ϴ��ļ�
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
					&error("�ϴ��ļ�ʧ�ܣ���ָ�����ļ����������޶����ļ���С60Kb!");
					exit 0;
				}
			} 
			close (OUTFILE);
			#print "�ļ����ϴ�";
			print "<script language='javascript'>window.close();</script>";
		}
	}
	else {
		&error("�ļ���ʽ����ȷ,���ϴ�*.jpg�ļ���*.gif�ļ���");
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
<title>����</title>
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
                          <div align="center" class="p9"><A href="javascript:history.back()"><font size=2>�� ��</font></a></div>
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
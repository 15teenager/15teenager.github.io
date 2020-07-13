#!d:/perl/bin/perl

####################################
$rootdir=&mypath;

####################################
$|++;						# Flush Output
#print "content-type:text/html\n\n";

my $dbpath = "$rootdir/db";
my $dbbakpath = "$rootdir/dbbackup";

opendir (DIR, $dbpath) or exit;
my @ls = readdir(DIR);
closedir (DIR);

my $file;
FILE: foreach $file (@ls) {
	# Skip the "." entry and ".." if we are at root level.
 	next FILE if  ($file eq '.');
	next FILE if  ($file eq '..');
	
	my $fullfile="$dbpath/$file";
	open (INFILE, "<$fullfile");
	binmode (INFILE);	# For those O/S that care.

	my $fullbakfile="$dbbakpath/$file";
	open(OUTFILE,">$fullbakfile");
	binmode (OUTFILE);	# For those O/S that care.
	
	while (read(INFILE, my $buffer, 1024)) { 
		print OUTFILE $buffer;
	}
	
	close INFILE;
	close(OUTFILE);
	
	#print "$fullfile\n<br>";
	#print "$fullbakfile\n<br>";
}

print "Location: manage.asp\n\n";




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
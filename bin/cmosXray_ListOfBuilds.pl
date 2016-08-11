#!/usr/bin/perl -w
use strict;
use warnings;
use Error qw(:try);

#use Win32::OLE qw/in with/;
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3; # Die on errors in Excel

use File::Basename;
use File::Copy;

use Time::Piece;
use Time::Seconds;

use Math::Round;


use Path::Class;

my $_debug =1;

###################################################
#
# cmosXray_ListOfBuilds.pl:	create a list of "model, serialNo, BuildLabel"
#							No input paramenter.
#							Output file: C:\\CMOSListOfBuild.txt
#			
###################################################


my @files;
### get x-ray directory
#dir('.')->recurse(callback => sub {
dir ("\\\\amersclnas02\\fire\\CMOS\\data")->recurse(callback => sub {
    my $file = shift;
	#if ($_debug>0) {print "$file\n";}
    if(lc($file) =~ /x-ray/) {
		my @subdirs=split(/\\/,$file);
		my $arrsize =@subdirs;
		next if (not lc($subdirs[$arrsize-1]) eq "x-ray"); #exit subroutine, so will not go further into subdir under x-ray directory.
		
		if ($_debug>0) {print "\t\t $file\n";}
        push @files, $file->absolute->stringify;
    }
	elsif (lc($file) =~ /setup/) {
		next;
	}
	elsif($file =~ /LED/){
		next;
	}
	elsif ($file =~ /ENG/) {
		next;
	}
	elsif ($file =~ /DXT1/) {
		next;
	}
	elsif ($file =~ /DFT/) {
		next;
	}
	elsif ($file =~ /CLT/) {
		next;
	}	
});





my $Excel;

my $recBook;
my $reportSheet;

### start excel application 
$Excel = Win32::OLE->GetActiveObject('Excel.Application')
   || Win32::OLE->new('Excel.Application', 'Quit');
$Excel -> {"Visible"} = 0;
$Excel -> {"DisplayAlerts"} = 0;  

my $txtSumFile ="C:\\CMOSListOfBuild.txt";
open(my $txtRecrodSum, ">$txtSumFile") or die "Can't open $txtSumFile for writing: $!\n";
### print header of col
print $txtRecrodSum "model,SeriNo,Build\n";


for my $file (@files) {
    #if ($_debug>0) {print "\n\n$file\n";}

	opendir(my $DH, $file) or die "Error opening $file: $!";
	my %xraydirs = map { $_ => (stat("$file/$_"))[10] } grep(! /^\.\.?$/, readdir($DH));
	closedir($DH);
	my @sorted_files = sort { $xraydirs{$b} <=> $xraydirs{$a} } (keys %xraydirs);
		
	if ($_debug>0) {
		for my $xrayfile (@sorted_files){
			print "\t\tsorted: $file/$xrayfile\n";
			
			my( $detmod, $detseri) =split(/[-.]/,$xrayfile);
			print "\t\t $detmod, $detseri\n";			
		}
	}
	
	my( $detmod, $detseri) =split(/[-.]/,$sorted_files[0]);
	
	next if (not( ($detmod eq "1207") or 
				  ($detmod eq "1512") or 
                  ($detmod eq "2307") or 
                  ($detmod eq "2315") or 
				  ($detmod eq "2923")) ) ;
	
	print "\n\n\nlast dir: $file/$sorted_files[0]\n";
	print "\t\t $detmod, $detseri\n";
    #my $num_test_recs =@sorted_files;
	#my $recDir = $file."\\".$sorted_files[$num_test_recs-1]."\\"."Test Record";
	my $recDir = $file."\\".$sorted_files[0]."\\"."Test Record";
	if (-d "$recDir") {
		# directory called cgi-bin exists
		print "record dir : $recDir, txtSumFile: $txtSumFile\n";
		#my $globpattern ="\"".$recDir."\\[*].xls*\"";
		my @reports;
		#my $num_reports =@reports;
		opendir(DIR, $recDir);
		@reports = grep(/\.xls$/,readdir(DIR));
		closedir(DIR);	
		print "\n\n\n";
		foreach my $theReport (@reports) {
			print "\t\t\t CofC ==>", $recDir."\\".$theReport,"\n";

			try {
				$recBook = $Excel->Workbooks->Open("\"".$recDir."\\".$theReport."\"");
				### activate CofC "Test Report" page   
				$reportSheet =$recBook->Worksheets("Test Report");
				$reportSheet->activate();
				
				my $cellname;
				$cellname ='F9';	print $txtRecrodSum "$detmod, $detseri,", $recBook->Worksheets("Test Report")->Range("$cellname:$cellname")->{value},",";
				
				print $txtRecrodSum "\n";
				$recBook->Close();	
			}			
			catch Error with {
			};
		}
		print "\n\n\n";
		
		
	}
	elsif (-e "$recDir") {
		# cgi-bin exists but is not a directory
		print "\"$recDir\" is not a directory!!!\n";
	}
	else {
		print "\"$recDir\" does not exist!!!\n";
	}	
}

close($txtRecrodSum);
$Excel->Quit();


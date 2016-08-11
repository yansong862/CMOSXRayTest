### updateLinearityData.pl usage:
###     input args: src directory, dest directory 
###     Notes: input data requirements:
###           one xlsx file under src directory, and one sheet named as "Linearity Calc", number of data columns are within "A" to "AM"
###           one xls file under dest directory, one sheet named as "Linearity", and one page named as "Test Report"


#!/usr/bin/perl -w
use strict;
use warnings;
use Win32::OLE qw/in with/;
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3; # Die on errors in Excel

use File::Basename;
use File::Copy;

use Time::Piece;
use Time::Seconds;

use Math::Round;

my $_debug =0;

# (1) quit unless we have the correct number of command-line args
my $num_args = $#ARGV + 1;
if ($num_args != 1) {
    print "\nUsage: printCofC.pl dest_dir\n";
    exit;
}
 
my($destDir) =@ARGV;
print ("\ninput args: $destDir\n");

 
my $Excel;
 
my $theDir;
my $destName;

my $srcBook;
my $destBook;
my $reportSheet;

###
###  change report logo and footer
###
### get dest file name
$theDir = $destDir; #"C:\\Users\\Public\\Documents\\testProg\\FaxitronCabinet\\Output\\Test Record";
opendir (DIR, $theDir) or die $!;

while (my $file = readdir(DIR )) {

    # Use a regular expression to ignore files beginning with a period
    next if ($file =~ m/^\./);

    # Use a regular expression to find files ending in .xls
    next unless ($file =~ m/\.xls$/);

	$destName =$theDir."\\".$file;
    print "test report file name:", "$destName\n";
};

closedir(DIR);



### start excel application 
$Excel = Win32::OLE->GetActiveObject('Excel.Application')
   || Win32::OLE->new('Excel.Application', 'Quit');
$Excel -> {"Visible"} = 0;
$Excel -> {"DisplayAlerts"} = 0;  

$destBook = $Excel->Workbooks->Open($destName);
### activate CofC "Test Report" page   
$reportSheet =$destBook->Worksheets("Test Report");
$reportSheet->activate();


my $bPrintCofC=1; #do not print if '0'
if ($bPrintCofC){
    print "Print CofC, if passed test.\n";

    $Excel = Win32::OLE->GetActiveObject('Excel.Application')
       || Win32::OLE->new('Excel.Application', 'Quit');
    $Excel -> {"Visible"} = 0;
    $Excel -> {"DisplayAlerts"} = 0;  

    $destBook = $Excel->Workbooks->Open($destName);

    $reportSheet =$destBook->Worksheets("Test Report");
    $reportSheet->activate();

	### get test result
	my $testResults =trim($destBook->Worksheets("Test Report")->Range("Y28:Y28")->{value});
	if ($testResults eq 'PASS') {
		print "test result: $testResults\n";
	
		$destBook->ActiveSheet()->PrintOut();
	}
    
    $destBook->Close();
    $Excel->Quit();
}

print "===>Done CofC Linearity data update.\n\n";





sub trim {
	my $value = $_[0];
	$value =~ s/^\s+//;
	$value =~ s/\s+$//;
	return $value;
}


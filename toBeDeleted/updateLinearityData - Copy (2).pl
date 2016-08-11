### updateLinearityData.pl usage:
###     input args: src directory and dest directory
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

# (1) quit unless we have the correct number of command-line args
my $num_args = $#ARGV + 1;
if ($num_args != 3) {
    print "\nUsage: updateLinearityData.pl source_dir, dest_dir, tempCSVFileName\n";
    exit;
}
 
# (2) we got two command line args, so assume they are the
# first name and last name
my $srcDir=$ARGV[0];
my $destDir=$ARGV[1];
my $tmpCSVFileName =$ARGV[2];
 
my $srcName;
my $destName;


### get src file name
my $theDir = $srcDir; #"C:\\Users\\Public\\Documents\\testProg\\FaxitronCabinet\\Linearity";

opendir (DIR, $theDir) or die $!;

while (my $file = readdir(DIR )) {

    # Use a regular expression to ignore files beginning with a period
    next if ($file =~ m/^\./);

    # Use a regular expression to find files ending in .xlsx
    next unless ($file =~ m/\.xlsx$/);

	$srcName =join("",$theDir,"\\", $file);
	print "$srcName\n";
};

closedir(DIR);


### get dest file name
$theDir = $destDir; #"C:\\Users\\Public\\Documents\\testProg\\FaxitronCabinet\\Output\\Test Record";
opendir (DIR, $theDir) or die $!;

while (my $file = readdir(DIR )) {

    # Use a regular expression to ignore files beginning with a period
    next if ($file =~ m/^\./);

    # Use a regular expression to find files ending in .xls
    next unless ($file =~ m/\.xls$/);

	$destName =$theDir."\\".$file;
    print "$destName\n";
};

closedir(DIR);


### found both src and dest file.
if (length($srcName)==0 || length($destName)==0) {
	print "'Linearity' or 'Test Record' file not found\n";
	exit(0);
}

### copy linearity data from light test linearity report to CofC "Linearity" sheet.
my $srcBook = $Excel->Workbooks->Open($srcName);
my $destBook = $Excel->Workbooks->Open($destName);
$srcBook->Worksheets("Linearity Calc")->Range("A:AM")->Copy;
$destBook->Worksheets("Linearity")->Range("A:AM")->PasteSpecial;

### activate CofC "Test Report" page   
my $reportSheet =$destBook->Worksheets("Test Report");
$reportSheet->activate();


### update footer
with ($reportSheet->PageSetup, 
    'RightFooter' => "&R&9Document: 68226\nRevison: 01");  

    

### update logo
my $bFoundPictures=0;

print "before delete shape count ", "Delete Op...\n";
my $picCount =$destBook->ActiveSheet->Shapes->count;
print $picCount,  "\n";
for (my $i=0; $i <$picCount; $i++) {
	my $shapeName =$destBook->ActiveSheet->Shapes($i+1)->Name;
	print $shapeName,"\n";
		
	if ($shapeName =~ /Picture/) {
		print "Found target\n";
		$bFoundPictures=1;
	}
}

# Dexela logo is "Picture 1". Therefore, only "Picture 1" will be deleted
if ($bFoundPictures) {
	print "Deleting....\n";
	$destBook->ActiveSheet->Shapes("Picture 1")->Delete;		
	#$destBook->ActiveSheet->Shapes(1)->Delete;		
}
print "after delete shape count ", $destBook->ActiveSheet->Shapes->count, "\n";

# insert PKI logo
print "Insert Op...\n";
#insertedALinkToPicture- my $picCurrent = $reportSheet->Pictures->Insert("C:\\CMOS\\ForReports\\PKILogo.png");
my $picCurrent = $reportSheet->Shapes->AddPicture("C:\\CMOS\\ForReports\\PKILogo.png", 0, 
					1, 225, 30, 90, 42);
$picCurrent->{Top} = 30;
$picCurrent->{Left} = 225;


### print 
#OnlyPrintPassed- $reportSheet-->ActiveSheet()->PrintOut();


$srcBook->Close(0);
$destBook->SaveAs({Filename =>$destName,FileFormat => xlOpenXMLWorkbook}); 

### make csv file name
# my $theCSVName;
# my $thePath;
# my $theExt;
# ($theCSVName,$thePath,$theExt) = fileparse($destName,".xls");
# $theCSVName =join("",$thePath,$theCSVName,".csv");
# $theCSVName =~ s/\\/\//g; #replace \ with /
# #print "CVS File Name:", $theCSVName, "\n";
# #$destBook->SaveAs({Filename =>$theCSVName,FileFormat => xlCSV}); 
print "temp CVS File name: ", $tmpCSVFileName, "\n";
$destBook->SaveAs({Filename =>$tmpCSVFileName,FileFormat => xlCSV}); 

$destBook->Close();
$Excel->Quit();




### get final Pass/Fail result 
my $bPrintCofC=0;

open(my $fh, "<", $tmpCSVFileName) 
	or die "cannot open ", $tmpCSVFileName, "\n"; 
while (my $fline = <$fh>) {
  chomp $fline;
  
  my $substring = uc("Overall result");
  my $strpass = "PASS";
  my $strfail = "FAIL";
    if (uc($fline) =~ /\Q$substring\E/) {
		print qq("$fline" contains "$substring"\n);
        ### in the report, "Fail" result will be printed in the next row. 
        ### Therefore, if "PASS" is not detected in this row, it will be considered as failed.
        if (uc($fline) =~ /\Q$strpass\E/) {
            $bPrintCofC=1;
            print "Pass!\n";
        }
        else {            
            print "Fail\n";
        }
	}
  #print "$fline\n";
}

close($fh);


if ($bPrintCofC){
    print "Print CofC for passed.\n";

    $Excel = Win32::OLE->GetActiveObject('Excel.Application')
       || Win32::OLE->new('Excel.Application', 'Quit');
    $Excel -> {"Visible"} = 0;
    $Excel -> {"DisplayAlerts"} = 0;  

    $destBook = $Excel->Workbooks->Open($destName);

    $reportSheet =$destBook->Worksheets("Test Report");
    $reportSheet->activate();
    #$reportSheet-->ActiveSheet()->PrintOut();
    
    $destBook->Close();
    $Excel->Quit();
}

print "===>Done CofC Linearity data update.\n\n";





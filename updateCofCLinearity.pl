### updateLinearityData.pl usage:
###     input args: 
###                 srcDir          - directory holds linearity data
###                 frmFile         - linearity data file
###                 destDir         - directory holds x-ray test record (CofC)
###                 iFrmFileType    - linearity data file type: 'txt' - text file; 'xl' -excel file. The script only checking for "txt". Otherwise, it will be considered as excel.
###
### examples:
### perl importLineartyData.pl C:\Users\scltester\Documents\GitWorkspace\FaxitronCabinet\Linearity linearity.txt "C:\Users\scltester\Documents\GitWorkspace\FaxitronCabinet\Output\Test Record" txt
### perl importLineartyData.pl C:\Users\scltester\Documents\GitWorkspace\FaxitronCabinet\Linearity "1512_11223 Test Record 22_1_2016.xls" "C:\Users\scltester\Documents\GitWorkspace\FaxitronCabinet\Output\Test Record" excel


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

my $_debug =1;

# (1) quit unless we have the correct number of command-line args
my $num_args = $#ARGV + 1;
if ($num_args != 4) {
    print "\nUsage: updateLinearityData.pl srcDir, frmFile, destSrc, iFrmFileType\n";   #iFromFileType: 1, from linearity text file, assume that the template is copied into the same directory. tab name "Linearity Calc" will hold the linearity data
                                                                                        #               2, from test record, tab name "Linearity"
    exit;
}
 
my ($srcDir, $frmFile, $destDir,$iFrmFileType) =@ARGV;

print "Input args: $srcDir, $frmFile, $destDir,$iFrmFileType\n";
 
my $Excel;
 
my $srcBook;
my $destBook;
my $reportSheet;

my ($srcName, $destName);
my $frmTabName;

###
###  update Linearity excel template
###

if ($iFrmFileType eq "txt") {

    my $linearityExcelTemp =join("",$srcDir, "\\LinearityData.xlsx"); #copied of: "C:\\CMOS\\ForReports\\Linearity Test_SerialNo.xlsx";
    print "linerity template: $linearityExcelTemp\n";

    #alreadyIncludeDirectoryName- my $inLinearityFile =join("",$srcDir,"\\", $frmFile); #example: "C:\\Users\\scltester\\Documents\\GitWorkspace\\FaxitronCabinet\\Linearity\\linearity.txt";
	my $inLinearityFile =$frmFile; #example: "C:\\Users\\scltester\\Documents\\GitWorkspace\\FaxitronCabinet\\Linearity\\linearity.txt";
    my $outLinearityFile ="C:\\CMOSLinearityCSVOutput.csv";

    ### remove the empty lines in txt file, save as csv file
    my @dataRows;
    open(my $outFH, '>', $outLinearityFile) or die "Could not open file '$outLinearityFile' $!";
    open(my $inFH, '<:encoding(UTF-8)', $inLinearityFile) or die "Could not open file '$inLinearityFile' $!";
    while (my $row = <$inFH>) {
      chomp $row;
      $row =~ s/\r//g; #remove carriage return from the line
      
      #print "length of the row:", length($row),"\n";
      if (length($row)>1) {
        print $outFH "$row\n";
        push @dataRows, $row;
      }
    }
    close $inFH;
    close $outFH;


    ### start excel application, copy text data into xlsx template
    $Excel = Win32::OLE->GetActiveObject('Excel.Application')
       || Win32::OLE->new('Excel.Application', 'Quit');
    $Excel -> {"Visible"} = 0;
    $Excel -> {"DisplayAlerts"} = 0;  
    my $linearityDataBook =$Excel->Workbooks->Open($outLinearityFile);
    $linearityDataBook->Worksheets("CMOSLinearityCSVOutput")->Range("A:AM")->Copy;

    my $xlBook = $Excel->Workbooks->Open($linearityExcelTemp); 
    my $linearSheet =$xlBook->Worksheets("Test File");
    $linearSheet->activate();

    $linearSheet->Range("A:AM")->PasteSpecial;

    $xlBook->SaveAs({Filename =>$linearityExcelTemp,FileFormat => xlOpenXMLWorkbook}); 

    $xlBook->Close();
    $linearityDataBook->Close();

    $Excel->Quit();

    $srcName =$linearityExcelTemp;
    $frmTabName ="Linearity Calc"; #the sheet in linearity data template
}
else {
    if (index($frmFile, 'Linearity') != -1) {
        $frmTabName ="Linearity Calc"; # get Linearity Data from Linearity excel file in Setup test.   
    }
    else {
        $frmTabName ="Linearity"; # get linearity data from Test Record.
    }
    $srcName =join("",$srcDir, "\\".$frmFile);
}

print "Linearity sheet name $frmTabName\n";



### get dest file name
my $theDir = $destDir; #"C:\\Users\\Public\\Documents\\testProg\\FaxitronCabinet\\Output\\Test Record";
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

print "get linearity data from $srcName to $destName\n";


### start excel application 
$Excel = Win32::OLE->GetActiveObject('Excel.Application')
   || Win32::OLE->new('Excel.Application', 'Quit');
$Excel -> {"Visible"} = 0;
$Excel -> {"DisplayAlerts"} = 0;  

### copy linearity data from light test linearity report to CofC "Linearity" sheet.
$srcBook = $Excel->Workbooks->Open($srcName);
$destBook = $Excel->Workbooks->Open($destName);
$srcBook->Worksheets($frmTabName)->Range("A:AM")->Copy;
$destBook->Worksheets("Linearity")->Range("A:AM")->PasteSpecial;


### activate CofC "Test Report" page   
$reportSheet =$destBook->Worksheets("Test Report");
$reportSheet->activate();

$srcBook->Close(0);
$destBook->SaveAs({Filename =>$destName,FileFormat => xlOpenXMLWorkbook}); 

$destBook->Close();
$Excel->Quit();



print "===>Done CofC Linearity data update.\n\n";





sub trim {
	my $value = $_[0];
	$value =~ s/^\s+//;
	$value =~ s/\s+$//;
	return $value;
}


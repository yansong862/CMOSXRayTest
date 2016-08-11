### updateLinearityData.pl usage:
###     input args: src directory, dest directory and tempCSVFileName
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
if ($num_args != 7) {
    print "\nUsage: updateLinearityData.pl source_dir, dest_dir, tempCSVFileName, theBuildName, bSlowMode, bNewGigENoiseSpec, bNewGigENoiseSpec\n";
    exit;
}
 
# (2) we got two command line args, so assume they are the
# first name and last name
my $srcDir=$ARGV[0];
my $destDir=$ARGV[1];
my $tmpCSVFileName =$ARGV[2];
my $theBuildName =uc($ARGV[3]);
my $bSlowMode =$ARGV[4]; #print "\n\n\n########## slowMode =$bSlowMode\n";
my $bMTFHighResolution =$ARGV[5];
my $bNewGigENoiseSpec =$ARGV[6];

print "source dir: $srcDir\n";
print "dest. dir: $destDir\n";
 
my $Excel;
 
my $theDir;
my $srcName;
my $destName;

my $srcBook;
my $destBook;
my $reportSheet;


my %detBuildInfo=();
my $buildInfofilename ='C:/CMOS/Configs/CMOSDetectorADCList.txt';
###
###  read 'C:/CMOS/Configs/CMOSDetectorADCList.txt' to get adc hareware info for CofC
###

open(FH, "<$buildInfofilename") or die "\nCan't open $buildInfofilename to read: $!\n";
    my @lines=<FH>;
	
    foreach my $theline (@lines) {
		next if (index($theline, "#") != -1); #skip comment line

		chomp($theline);
		
		next if (length($theline)==0);
		
        #my ($buildLabel, $buildRev, $ADCType,$ADCRev,$DAQRev)=split(/:*\s+/, $theline);
		my @words=split(/[;,\s]+/, $theline);
		next if (@words<5);
        #"$buildLabel, \t$buildRev, \t$ADCType, \t$ADCRev, \t$DAQRev\n";
		$words[0] =uc(trim($words[0]));
		print "$words[0], \t$words[1], \t$words[2], \t$words[3], \t$words[4]\n";
		
		$detBuildInfo{$words[0]} =join("", $words[1],",",$words[2],",",$words[3],",",$words[4]);
    }
close(FH);

if ($_debug) {
	print "\n\n";
	foreach my $key (keys %detBuildInfo ){
		my @values =split(/,/,$detBuildInfo{$key});
		print "$key, $values[0], $values[1], $values[2], $values[3]\n";
		if ($key eq $theBuildName) {
			print "\t\t !!!!!!!! found match!\n";
		}		
	}
}


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


### update some cell values
my $time = localtime;
#noNeed- $time -= ONE_DAY;
$destBook->Worksheets("Test Report")->Range("F16:F16")->{value} =[$time->strftime(' %m/%d/%Y')];

### update build rev, hardware rev
#replacedByInputArg- my $theBuildName =$destBook->Worksheets("Test Report")->Range("F9:F9")->{value};
if ($theBuildName eq "1207N-C16-HECA-24V") {
    $theBuildName ="1207N-C16-HECA";
}
if ($theBuildName eq "1512N-C16-HECA-24V") {
    $theBuildName ="1512N-C16-HECA";
}
$destBook->Worksheets("Test Report")->Range("F9:G9")->{value} =$theBuildName;
my $serNo =$destBook->Worksheets("Test Report")->Range("F10:F10")->{value};
foreach my $key (keys %detBuildInfo ){
	my @values =split(/[;,\s]+/,$detBuildInfo{$key});
	next if (@values<4);
	
	if ($_debug) {
		print "\n$key, $values[0], $values[1], $values[2], $values[3]\n";
		print "$theBuildName\n";
		print "$key\n";
	}
	
	next if (not $key eq $theBuildName);
    #next if (index(uc($theBuildName), uc($key)) == -1); #need use the above 'eq' check.
    
    print "\n******** Match $key, $values[0], $values[1], $values[2], $values[3]\n";
	
	### build rev
	$destBook->Worksheets("Test Report")->Range("F12:G12")->{value} =$values[0];
	$destBook->Worksheets("Test Report")->Range("F13:G13")->{value} =$values[2];
	$destBook->Worksheets("Test Report")->Range("F14:G14")->{value} =$values[3];
	
	#if ( ($key eq "1512N-C16-DRZS") or ($key eq "1512N-C16-HRCC") ) { #slow mode, only bin1x1 available
    if ( (uc($bSlowMode) eq "TRUE") ) { #slow mode, only bin1x1 available
		$destBook->Worksheets("Test Report")->Range("O10:O10")->{value} ="na";
		$destBook->Worksheets("Test Report")->Range("O11:O11")->{value} ="na";
		$destBook->Worksheets("Test Report")->Range("O12:O12")->{value} ="na";
		$destBook->Worksheets("Test Report")->Range("O13:O13")->{value} ="na";

		$destBook->Worksheets("Test Report")->Range("AA22:AA22")->{value} ="na";
		$destBook->Worksheets("Test Report")->Range("AA23:AA23")->{value} ="na";
		$destBook->Worksheets("Test Report")->Range("AA24:AA24")->{value} ="na";
		$destBook->Worksheets("Test Report")->Range("AA25:AA25")->{value} ="na";
        
	}

    if ((uc($bMTFHighResolution) eq "TRUE") ) { #High Resolution
        ###High Resolution
        ### update the spec and update formula  
        $destBook->Worksheets("MTF")->activate();    
        $destBook->Worksheets("MTF")->Range("M28:M28")->{value} =">0.600";
        $destBook->Worksheets("MTF")->Range("M29:M29")->{value} =">0.350";
        $destBook->Worksheets("MTF")->Range("M30:M30")->{value} =">0.200";
        
        $destBook->Worksheets("MTF")->Range("O28:O28")->{value} ="=IF(AND(ROUND(N28,3)>0.600,N28<>\"\"),\"PASS\",\"FAIL\")";
        $destBook->Worksheets("MTF")->Range("O29:O29")->{value} ="=IF(AND(ROUND(N29,3)>0.350,N29<>\"\"),\"PASS\",\"FAIL\")";
        $destBook->Worksheets("MTF")->Range("O30:O30")->{value} ="=IF(AND(ROUND(N30,3)>0.200,N30<>\"\"),\"PASS\",\"FAIL\")";
    
        $reportSheet =$destBook->Worksheets("Test Report");
        $reportSheet->activate();       
        $destBook->Worksheets("Test Report")->Range("M40:M40")->{value} =">0.600";
        $destBook->Worksheets("Test Report")->Range("M41:M41")->{value} =">0.350";
        $destBook->Worksheets("Test Report")->Range("M42:M42")->{value} =">0.200";
    }
    else {
        ###Low Resolution
        ### update the spec and update formula  
        $destBook->Worksheets("MTF")->activate();    
        $destBook->Worksheets("MTF")->Range("M28:M28")->{value} =">0.450";
        $destBook->Worksheets("MTF")->Range("M29:M29")->{value} =">0.150";
        $destBook->Worksheets("MTF")->Range("M30:M30")->{value} =">0.070";
        
        $destBook->Worksheets("MTF")->Range("O28:O28")->{value} ="=IF(AND(ROUND(N28,3)>0.450,N28<>\"\"),\"PASS\",\"FAIL\")";
        $destBook->Worksheets("MTF")->Range("O29:O29")->{value} ="=IF(AND(ROUND(N29,3)>0.150,N29<>\"\"),\"PASS\",\"FAIL\")";
        $destBook->Worksheets("MTF")->Range("O30:O30")->{value} ="=IF(AND(ROUND(N30,3)>0.070,N30<>\"\"),\"PASS\",\"FAIL\")";

        $reportSheet =$destBook->Worksheets("Test Report");
        $reportSheet->activate();       
        $destBook->Worksheets("Test Report")->Range("M40:M40")->{value} =">0.450";
        $destBook->Worksheets("Test Report")->Range("M41:M41")->{value} =">0.150";
        $destBook->Worksheets("Test Report")->Range("M42:M42")->{value} =">0.070";
    }
    
    if ((uc($bNewGigENoiseSpec) eq "TRUE") ) { #new GigE noise spec: 8.5 instead of 7.5adu
        $destBook->Worksheets("System Noise")->activate();    
        $destBook->Worksheets("System Noise")->Range("E3:E3")->{value} ="< 8.50 adu";
        $destBook->Worksheets("System Noise")->Range("E5:E5")->{value} ="< 8.50 adu";
        $destBook->Worksheets("System Noise")->Range("E7:E7")->{value} ="< 8.50 adu";
        
        $destBook->Worksheets("System Noise")->Range("G3:G3")->{value} ="=IF(AND(F3<8.5,F3>0,F3<>\"\"),\"PASS\",\"FAIL\")";
        $destBook->Worksheets("System Noise")->Range("G5:G5")->{value} ="=IF(AND(F5<8.5,F5>0,F5<>\"\"),\"PASS\",\"FAIL\")";
        $destBook->Worksheets("System Noise")->Range("G7:G7")->{value} ="=IF(AND(F7<8.5,F7>0,F7<>\"\"),\"PASS\",\"FAIL\")";
        
        $reportSheet =$destBook->Worksheets("Test Report");
        $reportSheet->activate();       
        $destBook->Worksheets("Test Report")->Range("M8:M8")->{value} ="< 8.50 adu";
        $destBook->Worksheets("Test Report")->Range("M10:M10")->{value} ="< 8.50 adu";
        $destBook->Worksheets("Test Report")->Range("M12:M12")->{value} ="< 8.50 adu";
    }
}

$destBook->Worksheets("Test Report")->Range("N16:N16")->{value} =["EM"];
$destBook->Worksheets("Test Report")->Range("N17:N17")->{value} =["1"];
$destBook->Worksheets("Test Report")->Range("X33:Z40")->{value} =[" "];

### set the serial number to 5-digit number
$destBook->Worksheets("Test Report")->Range("F10:G10")->{NumberFormat} ="00000";

### check left footer to force 5-digit serial number
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- my $newLFooter;
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- for my $position (qw(Header Footer)){
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 	for my $element (qw(Left Center Right)){
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 		my $item =$element.$position;
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 		if ($element.$position eq "LeftFooter") {
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 			my @LFooters=split(/-/,$destBook->Worksheets("Test Report")->PageSetup->$item);
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 			my $LFooterSize =@LFooters;
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 			#print "$LFooters[$LFooterSize-1]\n";
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 			
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 			$newLFooter=$LFooters[0];
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 			for (my $i =1; $i<$LFooterSize-1; $i++) {
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 				$newLFooter =join("", $newLFooter,"-",$LFooters[$i]);
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 			}
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 			$newLFooter =join("", $newLFooter,"-", sprintf("%05d", $LFooters[$LFooterSize-1]));
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 			print "New L.Footer in Report: $newLFooter\n";
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 			#print "$element$position: ", $Book3->Worksheets("Test Report")->PageSetup->$item, "\n";
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 		}
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 	}
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- }
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- 
#usePassedInParamter"$theBuildName"WillNotNeedDoThis- $destBook->Worksheets("Test Report") -> PageSetup -> {LeftFooter}   = $newLFooter; #"Left\nFooter";
$destBook->Worksheets("Test Report") -> PageSetup -> {LeftFooter}   = $theBuildName."-".$serNo; #"Left\nFooter";


$destBook->SaveAs({Filename =>$destName,FileFormat => xlOpenXMLWorkbook}); 

$destBook->Close();
$Excel->Quit();







###
###  update Linearity excel template
###

### copy linearity excel template to local
my $linearityExcelTemp =join("",$srcDir, "\\LinearityData.xlsx"); #"C:\\CMOS\\ForReports\\Linearity Test_SerialNo.xlsx";
print "linerity template: $linearityExcelTemp\n";

my $inLinearityFile =join("",$srcDir,"\\", "linearity.txt"); #"C:\\Users\\scltester\\Documents\\GitWorkspace\\FaxitronCabinet\\Linearity\\linearity.txt";
my $outLinearityFile ="C:\\CMOSLinearityCSVOutput.csv";

### remove the empty line
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


### start excel application 
$Excel = Win32::OLE->GetActiveObject('Excel.Application')
   || Win32::OLE->new('Excel.Application', 'Quit');
$Excel -> {"Visible"} = 0;
$Excel -> {"DisplayAlerts"} = 0;  
my $linearityDataBook =$Excel->Workbooks->Open($outLinearityFile);
$linearityDataBook->Worksheets("CMOSLinearityCSVOutput")->Range("A:AM")->Copy;

my $xlBook = $Excel->Workbooks->Open($linearityExcelTemp); 
my $linearSheet =$xlBook->Worksheets("Test File");
$linearSheet->activate();


#$linearSheet->write_string(0, 0, "test");
$linearSheet->Range("A:AM")->PasteSpecial;

$xlBook->SaveAs({Filename =>$linearityExcelTemp,FileFormat => xlOpenXMLWorkbook}); 

$xlBook->Close();
$linearityDataBook->Close();


$Excel->Quit();







###
###  update linearity data in final report
###

### get src file name
#useLinearitytemplateForNow- $theDir = $srcDir; #"C:\\Users\\Public\\Documents\\testProg\\FaxitronCabinet\\Linearity";
#useLinearitytemplateForNow- 
#useLinearitytemplateForNow- opendir (DIR, $theDir) or die $!;
#useLinearitytemplateForNow- 
#useLinearitytemplateForNow- while (my $file = readdir(DIR )) {
#useLinearitytemplateForNow- 
#useLinearitytemplateForNow-     # Use a regular expression to ignore files beginning with a period
#useLinearitytemplateForNow-     next if ($file =~ m/^\./);
#useLinearitytemplateForNow- 
#useLinearitytemplateForNow-     # Use a regular expression to find files ending in .xlsx
#useLinearitytemplateForNow- #useLinearitytemplateForNow-     next unless ($file =~ m/\.xlsx$/);
#useLinearitytemplateForNow- 
#useLinearitytemplateForNow- #useLinearitytemplateForNow- 	$srcName =join("",$theDir,"\\", $file);
#useLinearitytemplateForNow- 	print "$srcName\n";
#useLinearitytemplateForNow- };
#useLinearitytemplateForNow- 
#useLinearitytemplateForNow- closedir(DIR);
$srcName =$linearityExcelTemp;

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

### start excel application 
$Excel = Win32::OLE->GetActiveObject('Excel.Application')
   || Win32::OLE->new('Excel.Application', 'Quit');
$Excel -> {"Visible"} = 0;
$Excel -> {"DisplayAlerts"} = 0;  

### copy linearity data from light test linearity report to CofC "Linearity" sheet.
$srcBook = $Excel->Workbooks->Open($srcName);
$destBook = $Excel->Workbooks->Open($destName);
$srcBook->Worksheets("Linearity Calc")->Range("A:AM")->Copy;
$destBook->Worksheets("Linearity")->Range("A:AM")->PasteSpecial;



### print system noise, MTF and Linearity Data into C:\CMOSTempLog.txt
my $CMOSTempLog ="C:\\CMOSTempLog.txt";

my $thecol1;

open(TempLog, ">$CMOSTempLog") or die "Can't open $CMOSTempLog for writing: $!\n";
print TempLog "#system_noise\n"	  ;

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J6:J6")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K6:K6")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L6:L6")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.1, $destBook->Worksheets("Test Report")->Range("N6:N6")->{value}), "\n";

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J7:J7")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K7:K7")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L7:L7")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.1, $destBook->Worksheets("Test Report")->Range("N7:N7")->{value}), "\n";

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J8:J8")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K8:K8")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L8:L8")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.1, $destBook->Worksheets("Test Report")->Range("N8:N8")->{value}), "\n";

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J9:J9")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K9:K9")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L9:L9")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.1, $destBook->Worksheets("Test Report")->Range("N9:N9")->{value}), "\n";

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J10:J10")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K10:K10")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L10:L10")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.1, $destBook->Worksheets("Test Report")->Range("N10:N10")->{value}), "\n";

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J11:J11")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K11:K11")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L11:L11")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.1, $destBook->Worksheets("Test Report")->Range("N11:N11")->{value}), "\n";

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J12:J12")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K12:K12")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L12:L12")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.1, $destBook->Worksheets("Test Report")->Range("N12:N12")->{value}), "\n";

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J13:J13")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K13:K13")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L13:L13")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.1, $destBook->Worksheets("Test Report")->Range("N13:N13")->{value}), "\n";

print TempLog "#MTF\n"	  ;;
$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J40:J40")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K40:K40")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L40:L40")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.1, $destBook->Worksheets("Test Report")->Range("N40:N40")->{value}), "\n";

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J41:J41")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K41:K41")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L41:L41")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.1, $destBook->Worksheets("Test Report")->Range("N41:N41")->{value}), "\n";

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J42:J42")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K42:K42")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L42:L42")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.1, $destBook->Worksheets("Test Report")->Range("N42:N42")->{value}), "\n";

### linearity data	  
my $xaxis;
my $yaxis;

print TempLog "#linearity\n"	  ;
$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G26:G26")->{value});
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L26:L26")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G27:G27")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L27:L27")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G28:G28")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L28:L28")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G29:G29")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L29:L29")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G30:G30")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L30:L30")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G31:G31")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L31:L31")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G32:G32")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L32:L32")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G33:G33")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L33:L33")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G34:G34")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L34:L34")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G35:G35")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L35:L35")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G36:G36")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L36:L36")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G37:G37")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L37:L37")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G38:G38")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L38:L38")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G39:G39")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L39:L39")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G40:G40")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L40:L40")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";

$xaxis =nearest(.001, $destBook->Worksheets("Linearity")->Range("G41:G41")->{value}), 
$yaxis =nearest(.000001, $destBook->Worksheets("Linearity")->Range("L41:L41")->{value});
$xaxis =~ s/\./p/g;
print TempLog "linearity_$xaxis\t$yaxis\n";
	  
close(TempLog);



### activate CofC "Test Report" page   
$reportSheet =$destBook->Worksheets("Test Report");
$reportSheet->activate();

$srcBook->Close(0);
$destBook->SaveAs({Filename =>$destName,FileFormat => xlOpenXMLWorkbook}); 

### make csv file name
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

#$bPrintCofC=0; #do not print until CofC fixed
if ($bPrintCofC){
    print "Print CofC for passed.\n";

    $Excel = Win32::OLE->GetActiveObject('Excel.Application')
       || Win32::OLE->new('Excel.Application', 'Quit');
    $Excel -> {"Visible"} = 0;
    $Excel -> {"DisplayAlerts"} = 0;  

    $destBook = $Excel->Workbooks->Open($destName);

    $reportSheet =$destBook->Worksheets("Test Report");
    $reportSheet->activate();
    $destBook->ActiveSheet()->PrintOut();
    
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


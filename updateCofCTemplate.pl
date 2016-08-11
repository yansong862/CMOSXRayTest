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

my $bNoBuildRev =1;

# (1) quit unless we have the correct number of command-line args
my $num_args = $#ARGV + 1;
if ($num_args != 10) {
    print "\nUsage: updateLinearityData.pl dest_dir, tmpCSVFileName, theBuildName, theBuildRev, theADCRev, theDAQRev, bSlowMode, bNewGigENoiseSpec, bNewGigENoiseSpec, bRMA\n";
    exit;
}
 
my($destDir,$tmpCSVFileName, $theBuildName, $theBuildRev, $theADCRev, $theDAQRev, $bSlowMode, $bMTFHighResolution, $bNewGigENoiseSpec, $bRMA) =@ARGV;
print ("\ninput args: $destDir,$tmpCSVFileName, $theBuildName, $theBuildRev, $theADCRev, $theDAQRev, $bSlowMode, $bMTFHighResolution, $bNewGigENoiseSpec\n");

 
my $Excel;
 
my $theDir;
my $srcName;
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


### update logo
my $bFoundPictures=0;

if ($_debug) {print "before delete shape count ", "Delete Op...\n";}
my $picCount =$destBook->ActiveSheet->Shapes->count;
if ($_debug) {print $picCount,  "\n";}
for (my $i=0; $i <$picCount; $i++) {
	my $shapeName =$destBook->ActiveSheet->Shapes($i+1)->Name;
	if($_debug) {print $shapeName,"\n";}
		
	if ($shapeName =~ /Picture/) {
		if ($_debug){print "Found target\n";}
		$bFoundPictures=1;
	}
}

# Dexela logo is "Picture 1". Therefore, only "Picture 1" will be deleted
if ($bFoundPictures) {
	if($_debug){print "Deleting....\n";}
	$destBook->ActiveSheet->Shapes("Picture 1")->Delete;		
	#$destBook->ActiveSheet->Shapes(1)->Delete;		
}
if($_debug){print "after delete shape count ", $destBook->ActiveSheet->Shapes->count, "\n";}

# insert PKI logo
if($_debug){print "Insert Op...\n";}
#insertedALinkToPicture- my $picCurrent = $reportSheet->Pictures->Insert("C:\\CMOS\\ForReports\\PKILogo.png");
my $picCurrent = $reportSheet->Shapes->AddPicture("C:\\CMOS\\ForReports\\PKILogo.png", 0, 
					1, 225, 30, 90, 42);
$picCurrent->{Top} = 30;
$picCurrent->{Left} = 225;


### update test time
my $time = localtime;
#noNeed- $time -= ONE_DAY;
$destBook->Worksheets("Test Report")->Range("F16:F16")->{value} =[$time->strftime(' %m/%d/%Y')];

my $serNo =$destBook->Worksheets("Test Report")->Range("F10:F10")->{value};
### update build rev, hardware rev
if ( (uc($bRMA) eq "TRUE") and ( ($serNo+0) < 30000) ) {
	if ( $theBuildName eq "1207N-C16-HECA-6V") {
		print "Changed build label from $theBuildName to ";
		$theBuildName = "1207N-C16-HECA";
		print "$theBuildName\n";
	}
}
else{
	if ( $theBuildName eq "1207N-C16-HECA-24V") {
		print "Changed build label from $theBuildName to ";
		$theBuildName = "1207N-C16-HECA";
		print "$theBuildName\n";
	}
	elsif ( $theBuildName eq "1512N-C16-HECA-24V") {
		print "Changed build label from $theBuildName to ";
		$theBuildName = "1512N-C16-HECA";
		print "$theBuildName\n";
	}
}

$destBook->Worksheets("Test Report")->Range("F9:G9")->{value} =$theBuildName;

$destBook->Worksheets("Test Report")->Range("F12:G12")->{value} =$theBuildRev;
if ($bNoBuildRev==1) {
	$destBook->Worksheets("Test Report")->Range("F12:G12")->{value} ="";
}
$destBook->Worksheets("Test Report")->Range("F13:G13")->{value} =$theADCRev;
$destBook->Worksheets("Test Report")->Range("F14:G14")->{value} =$theDAQRev;


### update binx22 and binx44 test data to 'na' in case of slow mode, which only have binx11 data
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


### update MTF criteria based High/Low resolution
if ((uc($bMTFHighResolution) eq "TRUE") ) { 
   ###High Resolution
   ### update the spec and update formula  
   $destBook->Worksheets("MTF")->activate();    
   $destBook->Worksheets("MTF")->Range("M28:M28")->{value} =">0.600";
   $destBook->Worksheets("MTF")->Range("M29:M29")->{value} =">0.350";
   $destBook->Worksheets("MTF")->Range("M30:M30")->{value} =">0.200";

   $destBook->Worksheets("MTF")->Range("O28:O28")->{value} ="=IF(AND(ROUND(N28,3)>0.600,N28<>\"\"),\"PASS\",\"FAIL\")";
   $destBook->Worksheets("MTF")->Range("O29:O29")->{value} ="=IF(AND(ROUND(N29,3)>0.350,N29<>\"\"),\"PASS\",\"FAIL\")";
   $destBook->Worksheets("MTF")->Range("O30:O30")->{value} ="=IF(AND(ROUND(N30,3)>0.200,N30<>\"\"),\"PASS\",\"FAIL\")";
    
    $destBook->Worksheets("MTF")->Range("O28:O28")->{value} ="=IF(AND(ROUND(N28,3)>0.600,N28<>\"\"),\"PASS\",\"FAIL\")";
    $destBook->Worksheets("MTF")->Range("O29:O29")->{value} ="=IF(AND(ROUND(N29,3)>0.350,N29<>\"\"),\"PASS\",\"FAIL\")";
    $destBook->Worksheets("MTF")->Range("O30:O30")->{value} ="=IF(AND(ROUND(N30,3)>0.200,N30<>\"\"),\"PASS\",\"FAIL\")";
    
    $destBook->Worksheets("MTF")->Range("P28:P28")->{value} ="=IF(AND(ROUND(N28,3)>0.600,N28<>\"\"),\"\",\"FAIL\")";
    $destBook->Worksheets("MTF")->Range("P29:P29")->{value} ="=IF(AND(ROUND(N29,3)>0.350,N29<>\"\"),\"\",\"FAIL\")";
    $destBook->Worksheets("MTF")->Range("P30:P30")->{value} ="=IF(AND(ROUND(N30,3)>0.200,N30<>\"\"),\"\",\"FAIL\")";
    
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
    
    $destBook->Worksheets("MTF")->Range("O28:O28")->{value} ="=IF(AND(ROUND(N28,3)>0.450,N28<>\"\"),\"PASS\",\"FAIL\")";
    $destBook->Worksheets("MTF")->Range("O29:O29")->{value} ="=IF(AND(ROUND(N29,3)>0.150,N29<>\"\"),\"PASS\",\"FAIL\")";
    $destBook->Worksheets("MTF")->Range("O30:O30")->{value} ="=IF(AND(ROUND(N30,3)>0.070,N30<>\"\"),\"PASS\",\"FAIL\")";
    
    $destBook->Worksheets("MTF")->Range("P28:P28")->{value} ="=IF(AND(ROUND(N28,3)>0.450,N28<>\"\"),\"\",\"FAIL\")";
    $destBook->Worksheets("MTF")->Range("P29:P29")->{value} ="=IF(AND(ROUND(N29,3)>0.150,N29<>\"\"),\"\",\"FAIL\")";
    $destBook->Worksheets("MTF")->Range("P30:P30")->{value} ="=IF(AND(ROUND(N30,3)>0.070,N30<>\"\"),\"\",\"FAIL\")";
    
    $reportSheet =$destBook->Worksheets("Test Report");
    $reportSheet->activate();       
    $destBook->Worksheets("Test Report")->Range("M40:M40")->{value} =">0.450";
    $destBook->Worksheets("Test Report")->Range("M41:M41")->{value} =">0.150";
    $destBook->Worksheets("Test Report")->Range("M42:M42")->{value} =">0.070";
}
    
### update noise criteria for GigE noiser board.	
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


### update MTF phantom info.
$destBook->Worksheets("Test Report")->Range("N16:N16")->{value} =["EM"];
$destBook->Worksheets("Test Report")->Range("N17:N17")->{value} =["1"];
$destBook->Worksheets("Test Report")->Range("X33:Z40")->{value} =[" "];

### set the serial number to 5-digit number
$destBook->Worksheets("Test Report")->Range("F10:G10")->{NumberFormat} ="00000";

### update footer
with ($reportSheet->PageSetup, 
    'RightFooter' => "&R&9Document: 68226\nRevison: 01");  
	
### update left footer
$destBook->Worksheets("Test Report") -> PageSetup -> {LeftFooter}   = $theBuildName."-".$serNo; #"Left\nFooter";


### save updates
$destBook->SaveAs({Filename =>$destName,FileFormat => xlOpenXMLWorkbook}); 

$destBook->Close();
$Excel->Quit();




$Excel = Win32::OLE->GetActiveObject('Excel.Application')
   || Win32::OLE->new('Excel.Application', 'Quit');
$Excel -> {"Visible"} = 0;
$Excel -> {"DisplayAlerts"} = 0;  

$destBook = $Excel->Workbooks->Open($destName);


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
print TempLog $thecol1, "\t", nearest(.001, $destBook->Worksheets("Test Report")->Range("N40:N40")->{value}), "\n";

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J41:J41")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K41:K41")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L41:L41")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.001, $destBook->Worksheets("Test Report")->Range("N41:N41")->{value}), "\n";

$thecol1 = join("", trim($destBook->Worksheets("Test Report")->Range("J42:J42")->{value}),"_", 
					trim($destBook->Worksheets("Test Report")->Range("K42:K42")->{value}),"_",
					trim($destBook->Worksheets("Test Report")->Range("L42:L42")->{value}));
$thecol1 =~ s/\s+/_/g;
print TempLog $thecol1, "\t", nearest(.001, $destBook->Worksheets("Test Report")->Range("N42:N42")->{value}), "\n";

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

$destBook->SaveAs({Filename =>$destName,FileFormat => xlOpenXMLWorkbook}); 


### make csv file name
print "temp CVS File name: ", $tmpCSVFileName, "\n";
$destBook->SaveAs({Filename =>$tmpCSVFileName,FileFormat => xlCSV}); 


$destBook->Close();
$Excel->Quit();


my $bPrintCofC=0; #do not print if '0'
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


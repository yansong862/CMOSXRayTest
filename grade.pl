use Win32::OLE;
Win32::OLE->Option(Warn =>3);
use Error qw(:try);
use POSIX qw(strftime);
use File::Find;
use File::Copy;
use File::Basename;
use strict; ## Be a good programmer, use strict.
use warnings;

use Tie::IxHash;

=head1 CMOSXRay_SQLReportGenerator.pl 
========================================================================
Ver		Date		Author		Remarks
========================================================================
1.0		06/24/2015	Yan Song	Initial release for CMOS XRay test grading.
                                Modified from CMOS_SQLReportGenerator Rev 1.0 due to the differenct usage of AMA
								Three input parameters: 
									detect model number
									detector serial number
									output image directory
========================================================================

=cut

# (1) quit unless we have the correct number of command-line args
my $num_args = $#ARGV + 1;
if ($num_args != 5) {
    print "\nUsage: grade.pl detector model number, detector serial number, rmaNo, buildLabel, output image dir\n";
    exit;
}


my ($age, $dir, $name);

sub trim {
	my $value = $_[0];
	$value =~ s/^\s+//;
	$value =~ s/\s+$//;
	return $value;
}
sub theLower{
	my ($grade, $finalGrade)=@_;
	if ($grade eq "Fail") {
		$finalGrade="Fail";
	}
	elsif ($grade eq "Standard" && $finalGrade ne "Fail"){
		$finalGrade="Standard";
	}
	elsif ($grade eq "Mammo" &&  $finalGrade !~ /Fail|Standard/i){
		$finalGrade="Mammo";
	}
	elsif ($grade eq "Platinum" && $finalGrade !~ /Fail|Standard|Mammo/i){
		$finalGrade="Platinum";
	}
	return $finalGrade;
}
sub latestDefectFile {
	return if ($_!~/Defect_Count.txt$|Defect_Count.txt\.gz$/i);
	return if defined $age and $age > (stat($_))[9];
	$age = (stat(_))[9];
	$name = $File::Find::name;
	$dir = $File::Find::dir;
	
	print "find file $name\n";
}

########################################################################
####                Common Objects
########################################################################

##########################################

#my $DetType=$ARGV[0];
my $panelTypeName=$ARGV[0];		#model
my $DutID=$ARGV[1];			#serial no
my $rmaNo =$ARGV[2];       #RMA No.
my $buildLabel =$ARGV[3];       #BOM Label
my $DefectGrade="C:\\CMOS\\ForReports\\DefectGrades.txt";
my $queuedir="\\\\amersclnas02\\Fire\\DBdata\\SQLReport\\CMOS\\queue";

undef $name;
undef $dir;

#my $dutDataDirectory = $AMA->GetDUTDirectory();
#my $dutAcqdImagesDirectory="C:\\Users\\scltester\\Documents\\GitWorkspace\\FaxitronCabinet\\Output"; #dirname(dirname($dutDataDirectory))."\\CLT";						#output\Defect Map directory??
my $dutAcqdImagesDirectory=$ARGV[4];
find(\&latestDefectFile, $dutAcqdImagesDirectory);											#1512-14975_Defect_Count.txt??

my $testesult=$name;
my $outTxt=$dir."\\$panelTypeName"."-".$DutID."_SQLReport.txt";
print "\n$testesult\n$outTxt";
print "\nSensor Model: $panelTypeName\n";

my	%hspecs=();
tie %hspecs, "Tie::IxHash";

# get spec hash
open(FH, "<$DefectGrade") or die "\nCan't opne $DefectGrade for reading : $!\n";
my @list=<FH>;
close(FH);

my %CMOSReport=();
tie %CMOSReport, "Tie::IxHash";
%CMOSReport=(
	"single_pixel"=>"single pixel",
	"small_cluster"=>"small cluster",
	"medium_cluster"=>"medium cluster",
	"large_cluster"=>"large cluster",
	"spot_defect"=>"spot defect",
	"single_column"=>"single column",
	"double_column"=>"double column",
	"single_row"=>"single row",
	"double_row"=>"double row"
);

foreach my $line (@list) {
	next if ($line=~/^\#|^\s+/);
	chomp($line);
	#my @specs($CTQ, $Superior, $FinePlus, $HighPlus, $StandardPlus)=split(/\s+/, $line);
	my ($grade, $model, @specs)=split(/\s+/, $line);
	next if ($model !~ $panelTypeName );
	my $i=0;
	foreach my $ctq (keys %CMOSReport) {
		push(@{$hspecs{$ctq}},$specs[$i]);
		$i++;
	}
}
#foreach my $k (keys %hspecs) {print "\n$k:\t", join(",\t", @{$hspecs{$k}}),"+"};

my $binmode;
my %CTQValue_1x1=();
my %CTQValue_2x2=();
my %CTQValue_4x4=();

tie %CTQValue_1x1, "Tie::IxHash";
tie %CTQValue_2x2, "Tie::IxHash";
tie %CTQValue_4x4, "Tie::IxHash";

if (-e $testesult) {
	open(FH, "<$testesult") or die "\nCan't open $testesult for reading: $!\n";
	my @list=<FH>;
	close(FH);
	foreach my $line (@list) {

		$binmode=1 if ($line=~/Defect_Map_1x1/i);
		$binmode=2 if ($line=~/Defect_Map_2x2/i);
		$binmode=4 if ($line=~/Defect_Map_4x4/i);
		next if ($line=~/^\s+|^\#/);
		next if ($line=~/^Image:/i);
		#next if ($line!~/\:/);
		#last if ($line=~/FinalGrade/i);
		chomp($line);
		my ($CTQ, $value)=split(/:\s*/,$line);
		#print "\n CTQ, value: $CTQ, $value-";
		if ($binmode==1) {
			$CTQValue_1x1{$CTQ}=$value;
		}
		elsif ($binmode==2) {
			$CTQValue_2x2{$CTQ}=$value;
		}
		elsif ($binmode==4) {
			$CTQValue_4x4{$CTQ}=$value;
		}
		#last if ($line=~/^spot defect:/i);
	}
}
else{
	print "cannot find $testesult";
	exit 1;
}
#foreach my $ke (keys %CTQValue) {print "\n$ke: ", $CTQValue{$ke}};

######### output the report ################
my $report_time=strftime("%H:%M:%S", localtime(time));
my $report_date=strftime("%m/%d/%Y", localtime(time));

my $TestHW=$ENV{COMPUTERNAME};
if ($TestHW=~/sclld414/i) {
	$TestHW="CLT1";
}
elsif ($TestHW=~/sclld417/i) {
	$TestHW="CLT2";
}
elsif ($TestHW=~/sclld412/i) {
	$TestHW="ZT3";
}
else {										
	$TestHW="ENG";
}

open(FH, ">$outTxt") or die "Can't open $outTxt for writing: $!\n";
print FH "PanelID \t", $DutID;
print FH "\nUUT_ID\t\t", $DutID;
print FH "\nPanel_Type\t", $panelTypeName;
print FH "\nADEPT_Version\t", "2.5.2";
print FH "\nProcess_Step\t", "d"; ### "s", image test; "d", detector test
print FH "\nTest_HW_Used\t", $TestHW;
print FH "\nOperator\t", $ENV{USERNAME};
print FH "\nReport_Time\t", $report_time;
print FH "\nReport_Date\t", $report_date;
print FH "\nTest_Mode\t", "Full";
#print FH "\ndetectorSignature\t", sprintf("0x%x",$detSig);
print FH "\ndetectorSignature\t", $buildLabel;
print FH "\nmControllerSignature\t", "NA";
print FH "\nRMA\t", $rmaNo;
print FH "\n\n";

my $finalGrade="Platinum";
foreach my $k (keys %CMOSReport) {
	my $CTQ=$CMOSReport{$k};
	my $nt=3-int((length($k)+4)/8);
	my $value=$CTQValue_1x1{$CTQ};
	$value=1 if ($value=~/Yes/i);
	$value=0 if ($value=~/No/i);
	if ($value=~/\./) {$value=sprintf("%.3f",$value)};
	my $grade;

	my $Standard=$hspecs{$k}[0];
	my $Mammo=$hspecs{$k}[1];
	my $Platinum=$hspecs{$k}[2];
	
	if ($Platinum eq "na" || $value<=$Platinum) {$grade="Platinum";}
	elsif ($Mammo eq "na" || $value<=$Mammo) {$grade="Mammo";}
	elsif ($Standard eq "na" || $value<=$Standard) {$grade="Standard";}
	else {$grade="Fail";}

	$finalGrade=theLower($grade, $finalGrade);
	print FH "\n$k", "_1x1", "\t"x$nt, $value, "\t"x1, $grade;
}
my $nt=3-int(length("FinalGrade_1x1")/8);
if ($finalGrade !~ /Fail/i) {
	print FH "\nFinalGrade_1x1", "\t"x$nt, "0", "\t"x1, $finalGrade;
}
else{
	print FH "\nFinalGrade_1x1", "\t"x$nt, "1", "\t"x1, $finalGrade;
}
print FH "\n#";
$finalGrade="Platinum";
foreach my $k (keys %CMOSReport) {
	my $CTQ=$CMOSReport{$k};
	my $nt=3-int((length($k)+4)/8);
	
	my $value=$CTQValue_2x2{$CTQ};
	$value=1 if ($value=~/Yes/i);
	$value=0 if ($value=~/No/i);
	if ($value=~/\./) {$value=sprintf("%.3f",$value)};
	my $grade;

	my $Standard=$hspecs{$k}[0];
	my $Mammo=$hspecs{$k}[1];
	my $Platinum=$hspecs{$k}[2];
	
	if ($Platinum eq "na" || $value<=$Platinum) {$grade="Platinum";}
	elsif ($Mammo eq "na" || $value<=$Mammo) {$grade="Mammo";}
	elsif ($Standard eq "na" || $value<=$Standard) {$grade="Standard";}
	else {$grade="Fail";}

	$finalGrade=theLower($grade, $finalGrade);
	print FH "\n$k", "_2x2", "\t"x$nt, $value, "\t"x1, $grade;
}
$nt=3-int(length("FinalGrade_2x2")/8);
if ($finalGrade !~ /Fail/i) {
	print FH "\nFinalGrade_2x2", "\t"x$nt, "0", "\t"x1, $finalGrade;
}
else{
	print FH "\nFinalGrade_2x2", "\t"x$nt, "1", "\t"x1, $finalGrade;
}
print FH "\n#";
$finalGrade="Platinum";
foreach my $k (keys %CMOSReport) {
	my $CTQ=$CMOSReport{$k};
	my $nt=3-int((length($k)+4)/8);
	
	my $value=$CTQValue_4x4{$CTQ};
	$value=1 if ($value=~/Yes/i);
	$value=0 if ($value=~/No/i);
	if ($value=~/\./) {$value=sprintf("%.3f",$value)};
	my $grade;

	my $Standard=$hspecs{$k}[0];
	my $Mammo=$hspecs{$k}[1];
	my $Platinum=$hspecs{$k}[2];
	
	if ($Platinum eq "na" || $value<=$Platinum) {$grade="Platinum";}
	elsif ($Mammo eq "na" || $value<=$Mammo) {$grade="Mammo";}
	elsif ($Standard eq "na" || $value<=$Standard) {$grade="Standard";}
	else {$grade="Fail";}

	$finalGrade=theLower($grade, $finalGrade);
	print FH "\n$k", "_4x4", "\t"x$nt, $value, "\t"x1, $grade;
}
$nt=3-int(length("FinalGrade_4x4")/8);
if ($finalGrade !~ /Fail/i) {
	print FH "\nFinalGrade_4x4", "\t"x$nt, "0", "\t"x1, $finalGrade;
}
else{
	print FH "\nFinalGrade_4x4", "\t"x$nt, "1", "\t"x1, $finalGrade;
}

print FH "\n";
close(FH);

print "done grading. ready to load to DB\n";

### add system noise, MTF, Linearity data to  grade file
my $CMOSTempLog ="C:\\CMOSTempLog.txt";
system("type $CMOSTempLog >>\"$outTxt\"");
print "cat $CMOSTempLog to $outTxt\n";



## copy test result to MS Access DB queue directory.
my $dated_file = basename($outTxt);
$report_time =~ s/:/_/g;
$report_date =~ s/\//-/g;
$dated_file =~ s/.txt/\_$report_date\_$report_time.txt/g;
my $queuefile =join("",$queuedir,"\\",$dated_file);
#copy($outTxt, $queuedir."\\".$dated_file) or print "Failed copying $outTxt to $queuedir\\$dated_file: $!\n";
copy($outTxt, $queuefile) or print "Failed copying $outTxt to $queuefile: $!\n";
print "copied $outTxt to $queuefile\n";


## Call SQL loader
#system("perl C:\\ADEPT\\Scripts\\SQLLoader_CMOS_Stage.pl $queuedir\\$dated_file");
system("perl SQLLoader_CMOS_Stage.pl $queuedir\\$dated_file");
print "Success to run SQL loader";


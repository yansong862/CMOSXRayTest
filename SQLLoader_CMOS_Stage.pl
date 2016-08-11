use POSIX qw(strftime);
use File::Basename;
use File::Copy;
use warnings;
use strict;
use Win32::ODBC;
use Tie::IxHash;

=head1 Perl script for data access to SQL server with Win32::ODBC.
-------------------------------------------------------------------------
Ver 	Date  		Author	Remarks
1.0 	06/24/2015	Yu Su	Original version for CMOS SQL stage loading
-------------------------------------------------------------------------

=cut #note: a blank line before =cut and =head1 is preferred

my $SQLreport=$ARGV[0];
my $DB;
my $archivedir;
my $logdir="\\\\amersclnas02\\Fire\\DBdata\\SQLReport\\CMOS\\log";

sub errorlog {
	my $msg=$_[0];
	my $report_time=strftime("%H_%M_%S", localtime(time));
	my $report_date=strftime("%m-%d-%Y", localtime(time));
	my $log=$logdir."\\".basename($SQLreport);
	#$log=~s/.txt/.$report_date\_$report_time.txt/;
	open (LOG, ">$log") or die "\nCannot open $log to write: $!\n";
	print $msg;
	print LOG $msg;
	close(LOG);
	exit();
}

# Get a fieldname hash, assume input is a clean formated data
open(FH, "<$SQLreport") or die "\nCan't open $SQLreport to read: $!\n";
my @lines=<FH>;
close(FH);
my %Report=();
tie %Report, "Tie::IxHash";
foreach my $line (@lines) {
	next if ($line=~/^\#|^\s+$/);
	chomp($line);
	my ($name, $value, $grade)=split(/:*\s+/, $line);
	push(@{$Report{$name}}, $value);
	push(@{$Report{$name}}, $grade); #$Report{$name}[0] is $value, $Report{$name}[1] is $grade, 
}

=pod

foreach my $key (keys %Report) {
	if (defined $Report{$key}[1]) {
		print "\n$key $Report{$key}[0] $Report{$key}[1]";}
	else {
		print "\n$key $Report{$key}[0]";
	}
}

=cut

=head1 skip now # exit if not a valid PROMIS compID use UUT_ID
my $compID=$Report{"PanelID"}[0]; # 2170-5
if ($compID !~ /\w+\d+/) {
	print "\nNot a valid compID: $compID\n";
	errorlog("\nNot a valid compID: $compID\n");
}

=cut

my $bUseProdDB =1; 

my $DSN; # = "pkiDSN";
#$DSN = "driver={SQL Server};server=OPTFREDEV02;database=DS_PANT;uid=scltester;pwd=adept01";
$DSN = "driver={SQL Server};server=OPTSCLSQL01;database=DS_PANT;uid=scltester;pwd=adept01";
if ($bUseProdDB ==0) {
	$DSN ="driver={SQL Server};server=OPTSCLSQL01;database=DR_DEV;uid=scltester;pwd=adept01";
	print "in test DB\n";
}


my $tableName;
my $pant_head="pant_head";
my $pant_data="pant_data";

if ($Report{"Panel_Type"}[0]=~/^1207|^1512|^2307|^2315|^2321|^2923/i) {
	$archivedir="\\\\amersclnas02\\Fire\\DBdata\\SQLReport\\CMOS\\archive_stage";
	$logdir=    "\\\\amersclnas02\\Fire\\DBdata\\SQLReport\\CMOS\\log_stage";
	#my $logdir="C:\\temp";
}
else{
	print "\nNot a recongnized paneltype: ".$Report{"Panel_Type"}[0]."\n";
	errorlog("\nNot a recongnized paneltype: ".$Report{"Panel_Type"}[0]."\n");
}

if (!($DB = new Win32::ODBC($DSN))){
	print "Failure opening connection. ".Win32::ODBC::Error(). "\n";
	errorlog("Failure opening connection. ".Win32::ODBC::Error(). "\n");
}
else {
	print "Success (connection #", $DB->Connection(), ")\n";
}

my @column;
my $columnStr;
my @values;
my $valueStr;
my $testTime;
# write to PANT_HEAD table
my $sqlStatement="SELECT * FROM $pant_head";
if(!$DB->Sql($sqlStatement)){
	@column=$DB->FieldNames();
}
else{
	my $err=$DB->Error;
	$DB->Close();
	print $err;
	errorlog($err);
}
foreach my $fname (@column) {
	if ($fname=~/Part/i){
		push @values, "'".$Report{"Panel_Type"}[0]."'";
	}
	elsif ($fname=~/testTime/i) {
		$testTime=$Report{"Report_Date"}[0]." ".$Report{"Report_Time"}[0];
		push @values, "'$testTime'"; 
	}
	elsif (defined $Report{$fname}[0]) {
		push @values, "'".$Report{$fname}[0]."'"; # this requires the SQl fieldname to be consistent with the report.
	}
	else{
		push @values, "'NA'";
	} 
}
#my $headStr=join("|", @column);
my $headStr="^".join("\$|^", @column)."\$";
$columnStr=join("],[", @column);
$columnStr="[".$columnStr."]";
$valueStr=join(",", @values);
#print "\n", $columnStr;
#print "\n", $valueStr;
print "\nheaderstr", $headStr;
if ($DB->Sql("INSERT INTO $pant_head ($columnStr) VALUES ($valueStr)")) {
	my $err=$DB->Error;
	$DB->Close();
	print $err;
	errorlog($err);
}

# write to PANT_DATA
$sqlStatement="SELECT * FROM $pant_data";
@column=();
@values=();
if(!$DB->Sql($sqlStatement)){
	@column=$DB->FieldNames();
}
else{
	my $err=$DB->Error;
	$DB->Close();
	print $err;
	errorlog($err);
}
$columnStr=join("],[", @column);
$columnStr="[".$columnStr."]";
#print "\n", $columnStr;
my $i=1;
foreach my $key (keys %Report) {
	next if ($key =~ /$headStr|UUT_ID|Panel_Type|Report_Time|Report_Date|RMA/i);
	if (defined $Report{$key}[1]) {
		my $grade=$Report{$key}[1];
		$grade=~s/\*//g;
		$valueStr="'".$Report{"PanelID"}[0]."',";
		$valueStr=$valueStr."'$testTime',";
		$valueStr=$valueStr."'".$i++."',";
		$valueStr=$valueStr."'".$key."',";
		$valueStr=$valueStr."'".$Report{$key}[0]."',";
		$valueStr=$valueStr."'$grade'";
		#print "\n$key $Report{$key}[0] $Report{$key}[1]";
		#print "\n", $valueStr;
	}
	else {
		$valueStr="'".$Report{"PanelID"}[0]."',";
		$valueStr=$valueStr."'$testTime',";
		$valueStr=$valueStr."'".$i++."',";
		$valueStr=$valueStr."'".$key."',";
		$valueStr=$valueStr."'".$Report{$key}[0]."',";
		$valueStr=$valueStr."'NA'";
		#print "\n", $valueStr;
	}
	if ($DB->Sql("INSERT INTO $pant_data ($columnStr) VALUES ($valueStr)")) {
		my $err=$DB->Error;
		$DB->Close();
		print $err;
		errorlog($err);
	}
}

$DB->Close();

# move $SQLreport to archive folder
copy($SQLreport, $archivedir) or die "Failed copying $SQLreport to $archivedir: $!\n";
unlink $SQLreport;

__END__

WOW!

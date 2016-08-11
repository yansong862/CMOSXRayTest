#!/usr/bin/perl -w
use strict;
use warnings;

use File::Basename;

use Time::Piece;
use Time::Seconds;

use constant xlDialogPrint => 8;

use Math::Round;

use POSIX;

my $dataLinesPerSensor =17;

# (1) quit unless we have the correct number of command-line args
my $num_args = $#ARGV + 1;
if ($num_args != 1) {
    print "\nUsage: prepLinearityTxt.pl local_linearity_data_dir\n";
    for (my $i=0; $i<$num_args; $i++) {
        print "arg[$i]: $ARGV[$i]\n";
    }
    exit;
}

my $localDir=$ARGV[0];

my $linearitydata =join("",$localDir,"\\", "linearity_org.txt");
my $linearitydata_Sensor1 =join("",$localDir,"\\", "linearity.txt");

system("copy $linearitydata_Sensor1 $linearitydata");


my @datalines=();
my $fh;

open($fh, "<$linearitydata") or die "Can't open $linearitydata for writing: $!\n";

my $rcnt=0;
while (my $row = <$fh>) {
  chomp $row;
  next if (length($row)<=1);
  
  $rcnt =$rcnt+1;
  #print "\t $rcnt ==> $row\n";
  
  push @datalines, $row;
}

close($fh);

my $numDataLines =@datalines;
my $numSensor =($numDataLines/$dataLinesPerSensor);
print "$numSensor sensor(s) -($numDataLines data lines)\n";
if ($numSensor <2) {
	print "1-sensor linearity data is ready.";
	die;
}

open( $fh, ">$linearitydata_Sensor1") or die "Can't open $linearitydata_Sensor1 for writing: $!\n";
	$rcnt=0;
	while($rcnt<$numDataLines) {
		print $fh "$datalines[$rcnt]\n";
		$rcnt=$rcnt+$numSensor;
	}
close($fh);

print "1-sensor linearity data is ready.\n";



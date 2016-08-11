#!/usr/bin/perl -w
use strict;
use warnings;
use Win32::OLE qw/in with/;
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3; # Die on errors in Excel

use File::Basename;
use File::Copy;

use File::Glob ':glob';

use Time::Piece;
use Time::Seconds;

### update some cell values
my $time = localtime;
#$time -= ONE_DAY;
my $thetime =[$time->strftime(' %m/%d/%Y')];
print $thetime->[0], "\n";


my $defectMapDir ="C:\\Users\\scltester\\Documents\\GitWorkspace\\FaxitronCabinet\\Output\\Defect Map\\";

my @defectMapFile = bsd_glob(join("",$defectMapDir,"\\*_Defect_Count.txt"));
my $numOfFile =@defectMapFile;
print $numOfFile,"\n";

if ($numOfFile==0) {
    die "No defect count file found.\n";
}

foreach my $file (@defectMapFile) {
    open(FH, "<$file") or die "\nCan't open $file to read: $!\n";
        my @lines=<FH>;
        foreach my $theline (@lines) {
            my ($name, $value, $grade)=split(/:*\s+/, $theline);
            print @lines,"\n";
        }
    close(FH);
}


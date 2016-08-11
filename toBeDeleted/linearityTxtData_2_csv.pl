#!/usr/bin/perl -w
use strict;
use warnings;


my $inLinearityFile ="C:\\Users\\scltester\\Documents\\GitWorkspace\\FaxitronCabinet\\Linearity\\linearity.txt";
my $outLinearityFile ="C:\\CMOSLinearityCSVOutput.csv";

open(my $outFH, '>', $outLinearityFile) or die "Could not open file '$outLinearityFile' $!";

open(my $inFH, '<:encoding(UTF-8)', $inLinearityFile) or die "Could not open file '$inLinearityFile' $!";
 
while (my $row = <$inFH>) {
  chomp $row;
  
  my @data =split (',', $row);
  
  my $num_data =@data;
  
  next if ($num_data<2);
  print $num_data, ":", @data,"\n";
  for (my $ii =0; $ii<6; $ii++){
    
    if ($ii<5) {
        print $outFH @data[$ii],"," ;
    }
    else {
    }
        print $outFH @data[$ii];
  }
}

close $outFH;


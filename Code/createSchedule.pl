#!/usr/bin/perl -w
use strict;

## NOTES:
##  This code was developed by Ryan Hernandez. Any questions? Reach out directly: ryan.hernandez@ucsf.edu.

## To do: 
##    


my $VERBOSE = 0; #set this variable  = 1 to print out verbose messages; primarily for debugging
my $DEBUG = 0;  #set this variable = 1 to print out lines to debug; mostly not useful!
my $nlines = 0; #variable for counting lines in files

sub readTSV {
  my @ARGV = @_;
  if(scalar(@ARGV) == 0){
    die "must include file name to subroutine read.tsv\n";
  }
  my $file = $ARGV[0];
  open(IN,"$file") or die "cannot read $file\n";
  my $data = ""; #this will store the whole file in a string
  my $nlines = 0;
  while(<IN>){ 
    $nlines++;
    $_ =~ s/[^ -~\t]//g; #get rid of all control characters
    if($_ eq "" || $_ !~ /[a-zA-Z0-9]/){
      print "skipping blank line $nlines\n" if($DEBUG);
    }
    chomp;
    $data = $data."\n".$_;
  }
  close(IN);
  print "read data file $file, with $nlines lines\n";

  ##    Now we need to collect cell entries that span multiple lines, and concatenate
  ##    This will work by iteratively finding quote-pairs, and removing tabs and newlines from them
  my $foundquote = 1;
  my $quoteRound = 0;
  while($foundquote == 1){
    $quoteRound++;
    if($data =~ /^([^"]*)"([^"]*)"(.*)/){
      if($DEBUG){
        print "breaks: ($-[0],$+[0]), ($-[1],$+[1]), ($-[2],$+[2]), ($-[3],$+[3])\n";
        print "-[0]=$-[0]:".substr($data, $-[0], 20)."...".substr($data, $+[0]-20, 20)."\n" if($+[0]>20);
        print "+[0]=$+[0]:".substr($data, $+[0], 50)."\n";
        print "-[1]=$-[1]:".substr($data, $-[1], 20)."...".substr($data, $+[1]-20, 20)."\n" if($+[1]>20);
        print "+[1]=$+[1]:".substr($data, $+[1], 50)."\n";
        print "-[2]=$-[2]:".substr($data, $-[2], 20)."...".substr($data, $+[2]-20, 20)."\n" if($+[2]>20);
        print "+[2]=$+[2]:".substr($data, $+[2], 50)."\n";
        print "-[3]=$-[3]:".substr($data, $-[3], 50)."\n";
        print "+[3]=$+[3]:".substr($data, $+[3], 50)."\n";
      }
      my $beg = $1; #data before start of quote
      my $inside = $2; #data inside the quotes
      my $end = substr($data, $+[2]+1); #data after the quotes

      $inside =~ s/[\n\t]/,/g; ## replace newlines and tabs with commas
      $inside =~ s/^\s+//g; # remove any leading spaces

      if($DEBUG){
        print "Round $quoteRound:\n";
        print "\t>$inside<\n\t>".substr($end,0,50),"<\n";
      }
      $data = $beg.$inside.$end;
    }
    else{
      $foundquote = 0;
    }
  }
  my @dataLinesSplit = split('\n',$data);
  my $ndataLinesSplit = scalar(@dataLinesSplit);
  print "found $ndataLinesSplit lines in $file\n" if($VERBOSE);
  return @dataLinesSplit;
}

## Set base directories for intput ($DIR) and output ($DIROUT):
##    In general, this code assumes all registration files will be in the directory:  $DIR/Data/
my $DIR = "../";
my $DIROUT = "../OUT";
if(!(-d $DIROUT)){
  `mkdir $DIROUT`;
}
## Set the start/stop times of each interview; this also sets the maximum number of interviews
my @TIMES = ();
$TIMES[0] = "9:45-10:00";
$TIMES[1] = "10:00-10:10";
$TIMES[2] = "10:15-10:25";
$TIMES[3] = "10:30-10:40";
$TIMES[4] = "11:15-11:25";
$TIMES[5] = "11:30-11:40";
$TIMES[6] = "11:45-11:55";
$TIMES[7] = "12:00-12:10";

my %Faculty = (); #faculty who registered; the keys of this hash are email addresses, and other info indexed in an array:
##  {email} ->
##  [0]=name;
##  [1]=zoom link;
##  [2]=total number of research interests;
##  [3-k]=each of their research interests;
my $nFaculty = 0; #total number of faculty registered
my %facultyNAMEtoEMAIL = (); #map faculty names to their email
my %facultyEMAILtoNAME = (); #map faculty emails to their name
my %FacAvail = (); #record whether faculty are available; default yes
my %cat2fac = (); #map of research categories to faculty
my %facCatPop = (); #popularity of categories;
my $nfacCatPop = 0; #total number of research categories selected across faculty
my %int2fac = (); #map of research interests to faculty
my %facIntPop = (); #popularity of interests;
my $nfacIntPop = 0; #total number of research interests selected across faculty

my %facCancel = (); #keep track of faculty that cancel to remove from schedule
my $nfacCancel = 0; #total number of faculty that have canceled
my %scholCancel = (); #keep track of scholars that cancel to remove from schedule
my $nscholCancel = 0; #total number of scholars that have canceled

my $facultyCancelFile = ""; #set this file once faculty have canceled their registration, eg: "$DIR/Data/Faculty_CANCEL.tsv";
if($facultyCancelFile ne ""){
  open(IN, $facultyCancelFile) or die "cannot read $facultyCancelFile\n";
  $nlines = 0;
  while(<IN>){
   $nlines++;
    if($nlines <= 1){ ## skip header line
      next;
    }
    chomp;
    $_ =~ s/[^ -~\t]//g; #get rid of all control characters
    my $name = lc $_; #store in lower case
    if(!exists($facCancel{$name})){
      $nfacCancel++;
      $facCancel{$name}++;
    }
  }
  print "there are $nfacCancel faculty cancellations\n";
}

my $scholarCancelFile = ""; #set this file once scholars have canceled their registration, eg: "$DIR/Data/Scholar_CANCEL.tsv";
if($scholarCancelFile ne ""){
  open(IN, "$DIR/Data/Scholar_CANCEL.tsv") or die "cannot read Scholar_CANCEL.tsv\n";
  $nlines = 0;
  while(<IN>){
    $nlines++;
    if($nlines <= 1){ ## skip header line
     next;
    }
    chomp;
    $_ =~ s/[^ -~\t]//g; #get rid of all control characters
    my $name = lc $_;
    if(!exists($scholCancel{$name})){
      $nscholCancel++;
     $scholCancel{$name}++;
    }
  }
  print "there are $nscholCancel scholar cancellations\n";
}

my $zoomlinkInReg = 1; #indicates if faculty zoom links are in registration file; if in a separate file set $zoomlinkInReg = 0
my %FacultyZoomLinks = (); #this stores all the faculty zoom links, with email addresses as keys
my $nFacultyZoomLinks = 0; #total number of zoom links
if($zoomlinkInReg == 0){ 
  ## now lets get faculty zoom links
  my $facZoomFile = "$DIR/Data/FacZoom.tsv";
  print "Looking for faculty zoom links in file $facZoomFile\n" if($VERBOSE);
  open(IN,"$facZoomFile") or die "cannot read $facZoomFile\n";
  $nlines = 0;
  while(<IN>){
    $nlines++;
    if($nlines <= 1){
      next;
    }
    chomp;
    $_ =~ s/[^ -~\t]//g; #get rid of all control characters
    my @sp = split('\t', $_);
    my $email = lc $sp[1];
    my $link = $sp[2];
    chomp($link);
    if(!exists($FacultyZoomLinks{$email})){
      $FacultyZoomLinks{$email} = $link;
      $nFacultyZoomLinks++;
    }
    else{
      print "error, $email duplicated in faculty zoom file?\n";
      exit;
    }
  }
}

## lets collect Faculty registration data
my $facRegFile = "$DIR/Data/Faculty_Registration.tsv";
my @facLinesSplit = readTSV($facRegFile);


$nlines = 0;
my @facRegHeaders = (); #this stores the faculty registration header information
my $nfacRegHeaders = 0;
my %facRegHeaderColumns = (); #this is a map of relevant headers to their column in the registration file
$facRegHeaderColumns{"FirstName"} = 17;
$facRegHeaderColumns{"LastName"} = 18;
$facRegHeaderColumns{"email"} = 19;
$facRegHeaderColumns{"ResearchCat"} = 20;
$facRegHeaderColumns{"ResearchInt1"} = 22;
$facRegHeaderColumns{"ResearchInt2"} = 24;
$facRegHeaderColumns{"ResearchInt3"} = 26;
$facRegHeaderColumns{"website"} = 29;
$facRegHeaderColumns{"zoomlink"} = 32;

my $maxColumnEntry = 0;
foreach my $key (keys %facRegHeaderColumns){
  if($facRegHeaderColumns{$key} > $maxColumnEntry){
    $maxColumnEntry = $facRegHeaderColumns{$key};
  }
}

foreach my $line (@facLinesSplit){
  $nlines++;
  chomp($line);
  $line =~ s/[^ -~\t]//g; #get rid of all control characters
  if($nlines == 3){ #from qualtrix, header info is on line 2... is yours the same?
    @facRegHeaders = split('\t', $line);
    $nfacRegHeaders = scalar(@facRegHeaders);
    if($VERBOSE){
      print "Check that the column names match the desired columns:\n";
      foreach my $col (sort {$facRegHeaderColumns{$a}<=> $facRegHeaderColumns{$b}} keys %facRegHeaderColumns){
       print "\t$col\[$facRegHeaderColumns{$col}\] =?= $facRegHeaders[$facRegHeaderColumns{$col}]\n";
      }
      print "\n";
    }
  }
  if($nlines <= 4){ #skip all other header lines
    next;
  }

  my @sp = split('\t', $line);
  if(scalar(@sp) == 1){
    next;
  }

  my $name = "$sp[$facRegHeaderColumns{\"FirstName\"}]$sp[$facRegHeaderColumns{\"LastName\"}]"; #capture the name: FirstLast
  my $CapName = "$sp[$facRegHeaderColumns{\"FirstName\"}] $sp[$facRegHeaderColumns{\"LastName\"}]";
  $CapName =~ s/[^a-zA-Z\s]//g; #this is the name as entered
  $name =~ s/[^a-zA-Z]//g; #strip non-alpha characters
  $name = lc $name; # cast to lowercase, this is easier to merge across data sets

  if(exists($facCancel{$name})){ #skip faculty who cancel
    next;
  }

  my $email = $sp[$facRegHeaderColumns{"email"}];
  $email =~ s/[^a-zA-Z0-9@\.-]//g;
  $email = lc $email; # cast to lowercase

  if(exists($Faculty{$email})){
    warn "Warning. $email repeated in registration file line $nlines; overwritting with new info\n";
  }

  ## if zoom link was not included in a separate document, collect it from registration data.
  if($zoomlinkInReg == 1){ 
    $FacultyZoomLinks{$email} = $sp[$facRegHeaderColumns{"zoomlink"}];
    $nFacultyZoomLinks++;
  }
  else{
    if(!exists($FacultyZoomLinks{$email})){
      warn "Did not find zoom link for $email in zoomlink file\n";
    } 
  } 

  ## now capture faculty research categories. They are comma delimited, so split them and add.
  my $numCat = 0;
  my @cat = ();
  my $resCat = lc $sp[$facRegHeaderColumns{"ResearchCat"}];
  $resCat =~ s/\s//g; #remove any white space
  @cat = split(',', $resCat);
  for(my $i=0; $i<scalar(@cat); $i++){
    $cat2fac{$cat[$i]}{$email} = $numCat; #add this to category to faculty map
    $facCatPop{$cat[$i]}++;
    $nfacCatPop++;
    $numCat++;
  }
  
  ## now capture faculty research interests.
  my $numInt = 0;
  my @int = ();
  
  for(my $i=1; $i<=3; $i++){
    $int[$i] = $sp[$facRegHeaderColumns{"ResearchInt$i"}];
    if($sp[$facRegHeaderColumns{"ResearchInt$i"}] =~ /Other/){ #this means there was "Other" research interest. Replace with user text
      $int[$i] = $sp[$facRegHeaderColumns{"ResearchInt$i"}+1];
    }

    $int[$i] =~ s/[^a-zA-Z0-9@\.-]//g; #strip extra non A-Z characters
    $int[$i] = lc $int[$i]; # make it lower case
    if($int[$i] ne "" && $int[$i] !~ /other/ && $int[$i] ne "nopreference"){
      $facIntPop{$int[$i]}++;
      $nfacIntPop++;
      $int2fac{$int[$i]}{$email} = $numInt;
      $numInt++;
    }
  }

  if(!exists($Faculty{$email})){ #now enter info into hash database
    $nFaculty++;
  }
  $Faculty{$email}[0] = $name; #lowercase name
  if(!exists($FacultyZoomLinks{$email})){
    die "faculty $email does not have a zoom link\n";
  }
  $Faculty{$email}[1] = $FacultyZoomLinks{$email};
  $Faculty{$email}[2] = $numInt;
  if($numInt == 0){
    warn "faculty $email has 0 interests\n";
  }
  else{
    for(my $i=1; $i<=$numInt; $i++){
     $Faculty{$email}[2+$i] = $int[$i];
    }
  }
  $facultyNAMEtoEMAIL{$name} = $email;
  $facultyEMAILtoNAME{$email} = $CapName;

  #set default faculty availability. updated below in choices
  for(my $i=0; $i<8; $i++){
    $FacAvail{$email}[$i] = 1;
  }

  if($VERBOSE){
    print "read faculty $CapName $email on line $nlines:\n";
    for(my $i=0; $i<scalar(@{$Faculty{$email}}); $i++){
      print "\t[$i]=$Faculty{$email}[$i]\n";
    }
  }
}
close(IN);
print "read $nFaculty unique faculty\n";

## now lets collect scholar registration data
my $scholRegFile = "$DIR/Data/Scholar_Registration.tsv";
my @scholLinesSplit = readTSV($scholRegFile);


#$ Ok, read scholar registration data into memory, now parse it!
##    Now we can split the data by lines and process the data!
my %Scholars = (); #scholars who registered
## {email} ->
##  [0]=name
##  [1]=num interests
##  [2-k]=interest k;
my $nScholars = 0;
my %scholN2E = (); #map scholar name to email
my %scholE2N = (); #map email to scholar name
my %cat2schol = (); #map of research categories to scholars
my %scholCatPop = (); #scholar research category popularity
my $nscholCatPop = 0;
my %int2schol = (); #map of research interests to schlars
my %scholIntPop = (); #scholar research interest popularity
my $nscholIntPop = 0;

my @scholRegHeaders = (); #this stores the scholar registration header information
my $nscholRegHeaders = 0;
my %scholRegHeaderColumns = (); #this is a map of relevant headers to their column in the registration file
$scholRegHeaderColumns{"Name"} = 0;
$scholRegHeaderColumns{"email"} = 1;
$scholRegHeaderColumns{"ResearchCat"} = 7;
$scholRegHeaderColumns{"ResearchInt1"} = 9;
$scholRegHeaderColumns{"ResearchInt2"} = 11;
$scholRegHeaderColumns{"ResearchInt3"} = 13;
$maxColumnEntry = 0;
foreach my $key (keys %scholRegHeaderColumns){
  if($scholRegHeaderColumns{$key} > $maxColumnEntry){
    $maxColumnEntry = $scholRegHeaderColumns{$key};
  }
}

$nlines = 0;
foreach my $line (@scholLinesSplit){
  if($line eq "" || $line !~ /[a-zA-Z0-9]/){
    next;
  }
  $nlines++;
  if($nlines == 1){ #Parse header line
    chomp($line);
    $line =~ s/[^ -~\t]//g; #get rid of all control characters
    @scholRegHeaders = split('\t', $line);
    $nscholRegHeaders = scalar(@scholRegHeaders);
    if($VERBOSE){
      print "\n\nNow check that the scolar column names match the desired columns:\n";
      foreach my $col (sort {$scholRegHeaderColumns{$a}<=> $scholRegHeaderColumns{$b}} keys %scholRegHeaderColumns){
       print "\t$col\[$scholRegHeaderColumns{$col}\] =?= $scholRegHeaders[$scholRegHeaderColumns{$col}]\n";
      }
      print "\n";
    }
    next;
  }

  my @sp = split('\t', $line);
  if(scalar(@sp) < @scholRegHeaders){
    die "check for newlines on line $nlines: $line\n";
  }
  my $name = $sp[$scholRegHeaderColumns{"Name"}];
  my $CapName = $name;
  $CapName =~ s/[^a-zA-Z\s]//g; #strip non-alpha characters
  $name =~ s/[^a-zA-Z]//g; #strip non-alpha characters
  $name = lc $name; # cast to lowercase
  if(exists($scholCancel{$name})){ # skip scholars who cancel
    print "scholar $name cancelled, skip reading data\n" if($DEBUG);
    next;
  }


  my $email = $sp[$scholRegHeaderColumns{"email"}];
  $email =~ s/[^a-zA-Z0-9@\.-]//g;
  $email = lc $email;

  my $numCat = 0;
  my @cat = ();
  my $resCat = lc $sp[$scholRegHeaderColumns{"ResearchCat"}];
  $resCat =~ s/[\s\"]//g; #remove whitespace and quotes
  @cat = split(",", $resCat);
  for(my $i=0; $i<scalar(@cat); $i++){
    $cat[$i] = lc $cat[$i];
    $scholCatPop{$cat[$i]}++;
    $nscholCatPop++;
    $cat2schol{$cat[$i]}{$email} = $numCat;
    $numCat++;
  }
  
  for(my $i=0; $i<scalar(@cat); $i++){
    print "$i: $cat[$i]\n" if($DEBUG);
  }

  my $numInt = 0;
  my @int = ();
  for(my $i=1; $i<=3; $i++){
    print "$i: ".($scholRegHeaderColumns{"ResearchInt$i"})." - ".($sp[$scholRegHeaderColumns{"ResearchInt$i"}])."\n" if($DEBUG);
    $int[$i] = $sp[$scholRegHeaderColumns{"ResearchInt$i"}];
    if($int[$i] =~ /Other/){ #this means there was "Other" research interest. Replace with user text
      $int[$i] = $sp[$scholRegHeaderColumns{"ResearchInt$i"}+1];
    }

    $int[$i] =~ s/[^a-zA-Z0-9@\.-]//g; #strip extra non A-Z characters
    $int[$i] = lc $int[$i]; # make it lower case
    if($int[$i] ne "" && $int[$i] !~ /other/ && $int[$i] ne "nopreference"){
      $scholIntPop{$int[$i]}++;
      $nscholIntPop++;
      $int2schol{$int[$i]}{$email} = $numInt;
      $numInt++;
    }
  }

  if(!exists($Scholars{$email})){
    $nScholars++;
    $Scholars{$email}[0] = $name;
    $Scholars{$email}[1] = $numInt;
    if($numInt == 0){
      print "scholar $email has no interests\n" if($VERBOSE);
    }
    else{
      for(my $i=0; $i<$numInt; $i++){
	      $Scholars{$email}[2+$i] = $int[$i+1];
      }
    }
    if($VERBOSE == 1){
      print "$name added: $nScholars\n";
    }
  }
  else{ #scholar email address present more than once...
    print "$email present multiple times (line $nlines)... overwriting previous data\n";
  }
  $scholN2E{$name} = $email;
  $scholE2N{$email} = $CapName;

  if($VERBOSE){
    print "read scholar $CapName $email on line $nlines:\n";
    for(my $i=0; $i<scalar(@{$Scholars{$email}}); $i++){
      print "\t[$i]=$Scholars{$email}[$i]\n";
    }
  }
}
close(IN);
print "read $nScholars unique scholars\n";

print "\npopularity of faculty vs scholar research categories:\n";
foreach my $cat (sort {$facCatPop{$b} <=> $facCatPop{$a}} keys %facCatPop){
  if($facCatPop{$cat} < 10){
    next;
  }
  if(!exists($scholCatPop{$cat})){
    $scholCatPop{$cat} = 0;
  }
  
  print sprintf "%40s\t%2d (%4f)\t%2d (%4f)\n",
    $cat,
    $facCatPop{$cat},
    $facCatPop{$cat}/$nfacCatPop,
    $scholCatPop{$cat},
    $scholCatPop{$cat}/$nscholCatPop;
}


print "\npopularity of faculty vs scholar research interests:\n";
foreach my $int (sort {$facIntPop{$b} <=> $facIntPop{$a}} keys %facIntPop){
  if($facIntPop{$int} < 10){
    next;
  }
  if(!exists($scholIntPop{$int})){
    $scholIntPop{$int} = 0;
  }
  
  print sprintf "%40s\t%2d (%4f)\t%2d (%4f)\n",
    $int,
    $facIntPop{$int},
    $facIntPop{$int}/$nfacIntPop,
    $scholIntPop{$int},
    $scholIntPop{$int}/$nscholIntPop;
}

## Now lets get faculty choices and availability
my %FacChoices = ();
my %nFacChoices = ();
my $nFacWithChoices = 0;
my %scholarCNT = (); # number of times scholar selected
my $nscholarCNT = 0; #number of unique scholars selected
my $totScholarCNT = 0; #total number of scholar selections
my $nNotAvail = 0; #number of slots faculty are not available
my $facChoiceFile = "$DIR/Data/Faculty_Choices.tsv";
my @facChoices = readTSV($facChoiceFile);

$nlines = 0;
foreach my $line (@facChoices){
  $nlines++;
  if($nlines <= 3){
    next;
  }

  print "facChoice $nlines: $line\n" if($DEBUG);

  my @sp = split('\t',$line);
  my $email = $sp[18];

  $email = lc $email; #cast to lowercase
  print "working on $email:\n" if($DEBUG);

  if(!exists($Faculty{$email})){
    if($VERBOSE == 1){
      print "did not find Faculty{$email}...\n";
      exit;
    }
    next;
  }
  
  if(!exists($nFacChoices{$email})){
    $nFacWithChoices++;
  }
  else{
    print "faculty $email entered choices twice... overwritting\n" if($VERBOSE);
  }
  $nFacChoices{$email} = 0;
  for(my $i=21; $i<=25; $i++){
    if(scalar(@sp) > $i){
      $sp[$i] =~ s/[^a-zA-Z]//g;
      $sp[$i] = lc $sp[$i];
      if($sp[$i] eq ""){
        next;
      }
      print "\tchose ".($i-20).": $sp[$i]\n" if($DEBUG);

      if(exists($scholN2E{$sp[$i]})){
        $FacChoices{$email}[$nFacChoices{$email}] = $scholN2E{$sp[$i]};
        $nFacChoices{$email}++;
        if(!exists($scholarCNT{$sp[$i]})){
	        $nscholarCNT++;
        }
        $scholarCNT{$sp[$i]}++;
        $totScholarCNT++;
      }
      else{
        if($VERBOSE == 1){
	        print "did not find :$sp[$i]: in scholars... ($email)\n";
        }
      }
    }
  }

  my $availSlots = 0;
  for(my $i=0; $i<=7; $i++){
    my $colid = 26+$i;
    if(scalar(@sp) > $colid){ #if no entry, assume Yes available
      $sp[$colid] =~ s/[^a-zA-Z]//g;
      if($sp[$colid] =~ /No/){
        $FacAvail{$email}[$i] = 0;
        $nNotAvail++;
      }
      else{
        $availSlots++;
      }
    }
    else{
      $availSlots++;
    }
    print "\tAvail($i): $FacAvail{$email}[$i]\n" if($DEBUG);
  }
  if($availSlots == 0){
    print "faculty $email not available at all...\n" if($VERBOSE);
  }
}
close(IN);

if($VERBOSE == 1){
  my $i=0;
  foreach my $s (sort {$scholarCNT{$b}<=>$scholarCNT{$a}} keys %scholarCNT){
    $i++;
    print "$i: $s -> $scholarCNT{$s}\n";
  }
}

print "$nFacWithChoices faculty chose $totScholarCNT scholars ($nscholarCNT unique) and are not available for $nNotAvail slots\n";

## now lets get scholar choices
my %ScholChoices = ();
my %nScholChoices = ();
my $nScholWithChoices = 0;
my %facultyCNT = (); # number of times faculty selected
my $nfacultyCNT = 0; #number of unique faculty selected
my $totFacultyCNT = 0; #total number of faculty selections
my $scholChoicesFile = "$DIR/Data/Scholar_Choices.tsv";
my @scholChoices = readTSV($scholChoicesFile);
$nlines = 0;
foreach my $line (@scholChoices){
  $nlines++;
  if($nlines <= 3){ #skip header
    next;
  }
  print "data line: $line\n" if($DEBUG);

  my @sp = split('\t',$line);
  my $email = $sp[18];
  $email = lc $email; #cast to lowercase
  print "working on $email:\n" if($DEBUG);

  if(!exists($Scholars{$email})){
    if($VERBOSE == 1){
      print "did not find Scholar{$email}... on line $nlines\n";
    }
    next;
  }

  if(!exists($nScholChoices{$email})){
    $nScholWithChoices++;
  }
  else{
    die "scholar $email in choices twice!";
  }

  $nScholChoices{$email} = 0;
  for(my $i=19; $i<=23 && $i<scalar(@sp); $i++){
    my $name = $sp[$i];
    $name =~ s/[^a-zA-Z]//g;
    $name = lc $name;
    print "\tchoice ".($i-18).": $name\n" if($DEBUG);

    if($name eq "" || exists($facCancel{$name})){ ## either no entry or chose a faculty who canceled.
      next;
    }

    if(exists($facultyNAMEtoEMAIL{$name}) && exists($Faculty{$facultyNAMEtoEMAIL{$name}})){
      $ScholChoices{$email}[$nScholChoices{$email}] = $facultyNAMEtoEMAIL{$name};
      $nScholChoices{$email}++;
      if(!exists($facultyCNT{$name})){
	      $nfacultyCNT++;
      }
      $facultyCNT{$name}++;
      $totFacultyCNT++;
    }
    else{
      print "did not find faculty :$name: in scholars interest list... ($email)\n";
      exit;
    }
  }
}
print "$nScholWithChoices scholars chose $totFacultyCNT faculty ($nfacultyCNT unique)\n";

if($VERBOSE == 1){
  my $i=0;
  foreach my $s (sort {$facultyCNT{$b}<=>$facultyCNT{$a}} keys %facultyCNT){
    $i++;
    print "$i: $s -> $facultyCNT{$s}\n";
  }
}

### NOW BUILD SCHEDULES

my $FAILED = 1;
my $TOTTRIES = 0;
my $MAXTRIES = 100;
my $MINScholInt = 4; #minimum interviews for each scholars
my $MINFacInt = 3; #minimum interviews for each faculty
my $MAXInt = 8; #maximum number of interview slots
my %FACsched = (); #faculty schedules
my %nFACsched = (); #number of interviews for each faculty
my %FACslots = (); # hash of open slots for faculty
my %nFACslots = (); #number of open slots for each faculty
my %SCHOLsched = ();#scholar schedules
my %nSCHOLsched = (); #number of interviews for each scholar
my %SCHOLslots = (); #hash of open slots for scholars
my %nSCHOLslots = (); #number of open slots for each scholar

while ($FAILED && $TOTTRIES < $MAXTRIES) {
  $FAILED = 0;
  $TOTTRIES++;
  if($VERBOSE == 1){
    print "Try number $TOTTRIES\n";
  }
  
  %FACsched = (); #faculty schedules
  %nFACsched = (); #number of interviews for each faculty
  %FACslots = (); # hash of open slots for faculty
  %nFACslots = (); #number of open slots for each faculty
  %SCHOLsched = ();#scholar schedules
  %nSCHOLsched = (); #number of interviews for each scholar
  %SCHOLslots = (); #hash of open slots for scholars
  %nSCHOLslots = (); #number of open slots for each scholar
  my %F2S = (); # pairing of faculty to scholars
  my %S2F = (); # pairing of scholars to faculty

  ## initialize faculty schedules
  my $nfacInit = 0;
  foreach my $fac (keys %Faculty){
    $nFACslots{$fac} = $MAXInt; #default to max available
    $nFACsched{$fac} = 0;
    for(my $i=0; $i<$MAXInt; $i++){
      if($FacAvail{$fac}[$i] == 0){
	      $FACsched{$fac}[$i] = "NA";
        $nFACsched{$fac}++;
	      $nFACslots{$fac}--; 
      }
      else{ #available!
	      $FACsched{$fac}[$i] = "";
	      $FACslots{$fac}{$i} = 1;
      }
    }
    $nfacInit++;
  }
  if($VERBOSE == 1){
    print "$nfacInit faculty schedules initialized\n";
  }

  ## initialize scholar schedules
  my $nscholInit = 0;
  foreach my $schol (keys %Scholars){
    $nSCHOLsched{$schol} = 0;
    $nSCHOLslots{$schol} = $MAXInt;
    for(my $i=0; $i<$MAXInt; $i++){
      $SCHOLsched{$schol}[$i] = "";
      $SCHOLslots{$schol}{$i} = 1;
    }
    $nscholInit++;
  }
  if($VERBOSE == 1){
    print "$nscholInit scholar schedules intitializes\n";
  }

  ## first start with faculty interests. Loop over faculty, pull out their first interest, then move on to second interest
  my %facmiss = ();
  my $nfacmiss = 0;
  my $totfacmiss = 0;
  for(my $i=0; $i<5; $i++){
    ## randomize faculty order
    my @facOrd = ();
    my %facOrd = ();
    my $nfacOrd = 0;
    while($nfacOrd < $nFaculty){
      my $fac = (keys %Faculty)[rand keys %Faculty];
      if(!exists($facOrd{$fac})){
	      $facOrd[$nfacOrd] = $fac;
	      $facOrd{$fac}++;
	      $nfacOrd++;
      }
    }

    for(my $faci=0; $faci<$nFaculty; $faci++){
      my $fac = $facOrd[$faci];
      if(!exists($nFacChoices{$fac})){
	      next;
      }
      if($i<$nFacChoices{$fac}){
	      my $schol = $FacChoices{$fac}[$i];
	      ##print "trying $i: $fac/$schol\n";
	      if($nSCHOLsched{$schol} < $MAXInt){
	        ## match! find random open slot
	        my %tryslots = ();
	        my $ntryslots = 0;
	        while($ntryslots<$nFACslots{$fac}){
	          my $try = (keys %{$FACslots{$fac}})[rand keys %{$FACslots{$fac}}];
	          ##print "\ttrying slot $try\n";
	          if(exists($tryslots{$try})){
	            ##print "tried that!\n";
	            next;
	          }
	          $tryslots{$try}++;
	          $ntryslots++;
            if(exists($SCHOLslots{$schol}{$try})){
              ##print "\tFound a slot! fill it in!\n";
              $FACsched{$fac}[$try] = $schol;
              $SCHOLsched{$schol}[$try] = $fac;
              delete($FACslots{$fac}{$try});
              $nFACslots{$fac}--;
              $nFACsched{$fac}++;
              delete($SCHOLslots{$schol}{$try});
              $nSCHOLslots{$schol}--;
              $nSCHOLsched{$schol}++;
              $F2S{$fac}{$schol}++;
              $S2F{$schol}{$fac}++;
              last;
            }
            else{
              ##print "doh, try again: SCHOLsched = $SCHOLsched{$schol}[$try]\n";
            }
          }
        }
        else{
          if($VERBOSE == 1){
            print "$schol cant match with $fac, nSCHOLsched=$nSCHOLsched{$schol}\n";
          }
          if(!exists($facmiss{$fac})){
            $nfacmiss++;
          }
          $facmiss{$fac}++;
          $totfacmiss++;
        }
      }
    }
  }
  if($VERBOSE == 1){
    print "$nfacmiss faculty didnt get a top choice, $totfacmiss overall\n";
  }

  ## Now fill in scholar interests in the same way
  my %scholmiss = ();
  my $nscholmiss = 0;
  my $totscholmiss = 0;
  my %scholFull = ();
  my $nscholFull = 0;
  foreach my $schol (keys %Scholars){
    if($nSCHOLslots{$schol}==0){
      if(!exists($scholFull{$schol})){
        $nscholFull++;
      }
      $scholFull{$schol}++;
    }
  }

  if($VERBOSE == 1){
    print "$nscholFull scholars schedules full by faculty choices!\n";
  }

  for(my $i=0; $i<5; $i++){
    ## randomize scholar order
    my @scholOrd = ();
    my %scholOrd = ();
    my $nscholOrd = 0;
    while($nscholOrd < $nScholars){
      my $schol = (keys %Scholars)[rand keys %Scholars];
      if(!exists($scholOrd{$schol})){
        $scholOrd[$nscholOrd] = $schol;
        $scholOrd{$schol}++;
        $nscholOrd++;
      }
    }

    for(my $scholi=0; $scholi<$nScholars; $scholi++){
      my $schol = $scholOrd[$scholi];
      ##print "$scholi:$schol\n";
      if(!exists($nScholChoices{$schol})){
        next;
      }
      if($nSCHOLslots{$schol}==0){
        if(!exists($scholFull{$schol})){
          $nscholFull++;
        }
        $scholFull{$schol}++;
      }

      ##print "nScholChoices{$schol} = $nScholChoices{$schol}\n";
      if($i<$nScholChoices{$schol}){
        my $fac = $ScholChoices{$schol}[$i];
        ##print ">$fac\n";
        if(exists($S2F{$schol}) && exists($S2F{$schol}{$fac})){
          next; # scholar/faculty already paired!
        }
        ##print "trying $i: $fac($nFACslots{$fac})/$schol\n";
        if($nFACslots{$fac} > 0){
          ## match! find random open slot
          my %tryslots = ();
          my $ntryslots = 0;
          while($ntryslots<$nFACslots{$fac}){
            my $try = (keys %{$FACslots{$fac}})[rand keys %{$FACslots{$fac}}];
            ##print "\ttrying slot $try\n";
            if(exists($tryslots{$try})){
              ##print "tried that!\n";
              next;
            }
            $tryslots{$try}++;
            $ntryslots++;
            if(exists($SCHOLslots{$schol}{$try})){
              ##print "\tFound a slot! fill it in!\n";
              $FACsched{$fac}[$try] = $schol;
              $SCHOLsched{$schol}[$try] = $fac;
              delete($FACslots{$fac}{$try});
              $nFACslots{$fac}--;
              $nFACsched{$fac}++;
              delete($SCHOLslots{$schol}{$try});
              $nSCHOLslots{$schol}--;
              $nSCHOLsched{$schol}++;
              $F2S{$fac}{$schol}++;
              $S2F{$schol}{$fac}++;

              ##print "pairing schol:$schol>fac:$fac\n";
              last;
            }
            else{
              ##print "doh, try again: SCHOLsched = $SCHOLsched{$schol}[$try]\n";
            }
          }
        }
        else{
          if($VERBOSE == 1){
            print "$schol cant match with $fac, nFACsched=$nFACsched{$fac}\n";
          }
          if(!exists($scholmiss{$schol})){
            $nscholmiss++;
          }
          $scholmiss{$schol}++;
          $totscholmiss++;
        }
      }
    }
  }
  if($VERBOSE == 1){
    print "$nscholmiss scholars didnt get a top choice, $totscholmiss overall:\n";
    foreach my $schol (sort keys %scholmiss){
      print "\t$schol\n";
    }
  }

  ##now make random matches!
  ## check how many interviews everyone got
  my @scholinthist = (); #scholar interview histogram
  my @facinthist = ();  #faculty interview histogram
  my @facfilledhist = (); 
  my $nfactoofew = 0;
  my %facWithTooFew = ();
  my $nscholtoofew = 0;
  for(my $i=0; $i<=$MAXInt; $i++){
    $scholinthist[$i] = 0;
    $facinthist[$i] = 0;
  }
  foreach my $fac (keys %Faculty){
    $facfilledhist[$MAXInt-$nFACslots{$fac}]++;
    $facinthist[$nFACsched{$fac}]++;
    if($nFACsched{$fac} < $MINFacInt){
      $facWithTooFew{$fac}++;
      $nfactoofew++;
    }
  }
  foreach my $schol (keys %Scholars){
    $scholinthist[$nSCHOLsched{$schol}]++;
    if($nSCHOLsched{$schol} < $MINScholInt){
      $nscholtoofew++;
    }
  }
  if($VERBOSE == 1){
    print "\nN\t#Fac\t#Schols\t#facFilled\n";
    for(my $i=0; $i<=$MAXInt; $i++){
      print "$i\t$facinthist[$i]\t$scholinthist[$i]\t$facfilledhist[$i]\n";
    }
    print "\n\n$nfactoofew faculty and $nscholtoofew scholars with too few interviews\n";
  }

  ## start with filling in faculty schedules
  my %facTrouble = ();
  my $nfacTrouble = 0;
  my $TRIES = 0;
  while($nfactoofew > 0 && $TRIES<1000){
    print "\ttrying again: nfactoofew=$nfactoofew; try=$TRIES\n" if($DEBUG);
    my $fac = (keys %facWithTooFew)[rand keys %facWithTooFew];
    if($nFACsched{$fac} == $MAXInt || exists($facTrouble{$fac})){
      $TRIES++;
      next;
    } 
    if($Faculty{$fac}[2] == 0){
      ##print "faculty $fac has no interests...\n";
      $facTrouble{$fac}++;
      $TRIES++;
      next;
    }
    if($VERBOSE == 1){
      print "trying fac=$fac ($nFACsched{$fac})\n";
    }

    ##now get list of scholars with overlapping interests
    my %schols = ();
    my $nschols = 0;
    for(my $i=0; $i<$Faculty{$fac}[2]; $i++){
      my $int = $Faculty{$fac}[3+$i];
      foreach my $schol (keys %{$int2schol{$int}}){
        if(exists($S2F{$schol}{$fac}) || $nSCHOLslots{$schol} == 0){
          next;
        }
        if(!exists($schols{$schol})){
          $schols{$schol}++;
          $nschols++;
        }
      }
    }
    if($VERBOSE == 1 && $nschols == 0){
      print "no scholars with overlapping interests ($Faculty{$fac}[2]) in:\n";
      for(my $i=0; $i<$Faculty{$fac}[2]; $i++){
        print "$Faculty{$fac}[3+$i]\n";
      }
      $facTrouble{$fac}++;
      $TRIES++;
      next;
    }
    ## draw a random scholar with overlapping interests
    my $schol = (keys %schols)[rand keys %schols];
    if(!defined($schol)){
      $TRIES++;
      next;
    }
    ##print "Found $nschols scholars with overlapping interest, trying $schol\n";
    
    my %tryslots = ();
    my $ntryslots = 0;
    while($ntryslots<$nFACslots{$fac}){
      my $try = (keys %{$FACslots{$fac}})[rand keys %{$FACslots{$fac}}];
      if($VERBOSE == 1){
        print "\ttrying schol=$schol slot $try\n";
      }

      if(exists($tryslots{$try})){
        ##print "tried that!\n";
        next;
      }
      $tryslots{$try}++;
      $ntryslots++;
      if(exists($SCHOLslots{$schol}) && exists($SCHOLslots{$schol}{$try})){
        ##print "\tFound a slot! fill it in!\n";
        $FACsched{$fac}[$try] = $schol;
        $SCHOLsched{$schol}[$try] = $fac;
        delete($FACslots{$fac}{$try});
        $nFACslots{$fac}--;
        $nFACsched{$fac}++;
        delete($SCHOLslots{$schol}{$try});
        $nSCHOLslots{$schol}--;
        $nSCHOLsched{$schol}++;
        $F2S{$fac}{$schol}++;
        $S2F{$schol}{$fac}++;
        last;
      }
      else{
        if($VERBOSE == 1){
          print "doh, try again: SCHOLsched = $SCHOLsched{$schol}[$try]\n";
        }
      }
    }
    
    $nfactoofew = 0;
    foreach my $fac (keys %Faculty){
      if($nFACsched{$fac} < $MINFacInt){
        $nfactoofew++;
        if($VERBOSE == 1){
          print "$fac has $nFACsched{$fac} scheduled $nFACslots{$fac} available\n";
        }
      }
    }
    if($VERBOSE == 1){
      print "try $TRIES: nfactoofew -> $nfactoofew\n";
    }
    $TRIES++;
  }

  ## now fill in random scholar schedules
  $nscholtoofew = 0;
  my %scholsWithTooFew = ();
  foreach my $schol (keys %Scholars){
    if($nSCHOLsched{$schol} < $MINScholInt){
      $scholsWithTooFew{$schol}++;
      $nscholtoofew++;
    }
  }
  print "scholars with too few interviews: $nscholtoofew\n" if($VERBOSE);
  
  my %scholTrouble = ();
  $TRIES = 0;
  while($nscholtoofew > 0 && $TRIES<10){
    my $schol = (keys %scholsWithTooFew)[rand keys %scholsWithTooFew];
    if($nSCHOLsched{$schol} == $MAXInt || exists($scholTrouble{$schol})){
      next;
    }
    if($Scholars{$schol}[1] == 0){
      ##print "scholar $schol has no interests...\n";
      $scholTrouble{$schol}++;
      $TRIES++;
    }
    if($VERBOSE == 1){
      print "trying schol=$schol ($nSCHOLsched{$schol})\n";
    }

    my %facs = ();
    my $nfacs = 0;
    for(my $i=0; $i<$Scholars{$schol}[1]; $i++){
      my $int = $Scholars{$schol}[2+$i];
      foreach my $fac (keys %{$int2fac{$int}}){
        if(exists($S2F{$schol}{$fac}) || $nFACslots{$fac} == 0){
          next;
        }
        if(!exists($facs{$fac})){
          $facs{$fac}++;
          $nfacs++;
        }
      }
    }
    my $fac = (keys %facs)[rand keys %facs];
    ##print "Found $nfacs faculty with overlapping interest, trying $fac\n";

    if(defined($fac) && $nFACslots{$fac} > 0){
      ## match! find random open slot
      my %tryslots = ();
      my $ntryslots = 0;
      while($ntryslots<$nSCHOLslots{$schol}){
        my $try = (keys %{$SCHOLslots{$schol}})[rand keys %{$SCHOLslots{$schol}}];
        ##print "\ttrying slot $try\n";
        if(exists($tryslots{$try})){
          ##print "tried that!\n";
          next;
        }
        $tryslots{$try}++;
        $ntryslots++;
        if(exists($FACslots{$fac}{$try})){
          ##print "\tFound a slot! fill it in!\n";
          $FACsched{$fac}[$try] = $schol;
          $SCHOLsched{$schol}[$try] = $fac;
          delete($FACslots{$fac}{$try});
          $nFACslots{$fac}--;
          $nFACsched{$fac}++;
          delete($SCHOLslots{$schol}{$try});
          $nSCHOLslots{$schol}--;
          $nSCHOLsched{$schol}++;
          $F2S{$fac}{$schol}++;
          $S2F{$schol}{$fac}++;
          last;
        }
        else{
          ##print "doh, try again: SCHOLsched = $SCHOLsched{$schol}[$try]\n";
        }
      }
    }
    $nscholtoofew = 0;
    foreach my $schol (keys %Scholars){
      if($nSCHOLsched{$schol} < $MINScholInt){
        $nscholtoofew++;
      }
    }
    ##print "nscholtoofew -> $nscholtoofew\n";
    $TRIES++;
  }

  ## now fill in rest of schedule
  for(my $sloti=0; $sloti<$MAXInt; $sloti++){
    ## get list of faculty/scholars available
    my %facAvail = ();
    my $nfacAvail = 0;
    foreach my $fac (keys %FACslots){
      if(exists($FACslots{$fac}{$sloti})){
        $facAvail{$fac}++;
        $nfacAvail++;
      }
    }
    my %scholAvail = ();
    my $nscholAvail = 0;
    foreach my $schol (keys %SCHOLslots){
      if(exists($SCHOLslots{$schol}{$sloti})){
        $scholAvail{$schol}++;
        $nscholAvail++;
      }
    }
    ##print "slot $sloti: nfacAvail=$nfacAvail; nscholAvail=$nscholAvail\n";
    foreach my $fac (sort {$nFACslots{$b} <=> $nFACslots{$a}} keys %facAvail){
      ##print "$sloti: $fac: $nFACslots{$fac}\n";
      my $found = 0;
      foreach my $schol (sort {$nSCHOLslots{$b} <=> $nSCHOLslots{$a}} keys %scholAvail){
        for(my $resint=0; $resint<$Faculty{$fac}[2]; $resint++){
          if(exists($int2schol{$Faculty{$fac}[3+$resint]}{$schol}) &&
            !exists($F2S{$fac}{$schol})){
            ## FOUND A MATCH!
            $FACsched{$fac}[$sloti] = $schol;
            $SCHOLsched{$schol}[$sloti] = $fac;
            delete($FACslots{$fac}{$sloti});
            $nFACslots{$fac}--;
            $nFACsched{$fac}++;
            delete($SCHOLslots{$schol}{$sloti});
            $nSCHOLslots{$schol}--;
            $nSCHOLsched{$schol}++;
            $F2S{$fac}{$schol}++;
            $S2F{$schol}{$fac}++;
            delete($scholAvail{$schol});
            $found = 1;
            last;
          }
        }
        if($found == 1){
          last;
        }
      }
    }
  }

  ## check how many interviews everyone got
  @scholinthist = ();
  @facinthist = ();
  my @facfillhist = ();
  my %factoofew = ();
  $nfactoofew = 0;
  my %scholtoofew = ();
  $nscholtoofew = 0;
  for(my $i=0; $i<=$MAXInt; $i++){
    $scholinthist[$i] = 0;
    $facinthist[$i] = 0;
    $facfillhist[$i] = 0;
  }
  foreach my $fac (keys %Faculty){
    $facinthist[$nFACsched{$fac}]++;
    my $empty = $MAXInt-$nFACslots{$fac};
    $facfillhist[$empty]++;
    if($nFACsched{$fac} < $MINFacInt){
      $factoofew{$fac}++;
      $nfactoofew++;
    }
  }
  foreach my $schol (keys %Scholars){
    $scholinthist[$nSCHOLsched{$schol}]++;
    if($nSCHOLsched{$schol} < $MINScholInt){
      $scholtoofew{$schol}++;
      $nscholtoofew++;
    }
  }

  if($VERBOSE == 1){
    print "\nN\t#FacInt\t#Schols\n";
    for(my $i=0; $i<=$MAXInt; $i++){
      print "$i\t$facfillhist[$i]\t$scholinthist[$i]\n";
    }

    print "\n\n$nfactoofew faculty and $nscholtoofew scholars with too few interviews\n";
    print "Faculty with too few interviews:\n";
    foreach my $fac (sort keys %factoofew){
      print "\t$fac\n";
    }
    print "Scholars with too few interviews:\n";
    foreach my $schol (sort keys %scholtoofew){
      print "\t$schol\n";
    }
  }

  #check if we are done...
  if(($nfactoofew>0 || $nscholtoofew>0) && $TOTTRIES < $MAXTRIES){
    $FAILED=1; #doh... try again!
  }
  else{ #done trying... declare victory!
    print "\n\n$nfactoofew faculty and $nscholtoofew scholars with too few interviews\n" if($VERBOSE);
    print "Faculty with too few interviews:\n" if($VERBOSE);
    foreach my $fac (sort keys %factoofew){
      print "\t$fac\n" if($VERBOSE);
    }
    print "Scholars with too few interviews:\n" if($VERBOSE);
    foreach my $schol (sort keys %scholtoofew){
      print "\t$schol\n" if($VERBOSE);
    }
    print "N\t#Faculty\t#Scholars\n" if($VERBOSE);
    for(my $i=0; $i<=$MAXInt; $i++){
      print "$i\t$facfillhist[$i]\t$scholinthist[$i]\n" if($VERBOSE);
    }

    ## now let's just fill in the schedule 
    print "done trying... now lets fill in more interviews completely randomly...\n" if($VERBOSE);
    my $minINTS = 0; #minimum number of interviews for any scholar/faculty...
    while($minINTS < $MINFacInt && $minINTS < $MINScholInt){
      $minINTS = 8;
      foreach my $schol (sort {$nSCHOLsched{$a} <=> $nSCHOLsched{$b}} keys %nSCHOLsched){
        if($nSCHOLsched{$schol} < $minINTS){
          $minINTS = $nSCHOLsched{$schol};
        }
        if($nSCHOLsched{$schol} > $MINScholInt){
          last;
        }
        foreach my $fac (sort {$nFACsched{$a} <=> $nFACsched{$b}} keys %nFACsched){
          if(exists($F2S{$fac}{$schol})){
            next;
          }
          if($nFACsched{$fac} == $MAXInt){
            last;
          }

          for(my $try=0; $try<$MAXInt; $try++){
            if(exists($SCHOLslots{$schol}{$try}) && exists($FACslots{$fac}{$try})){
              ##print "\tFound a slot! fill it in!\n";
              $FACsched{$fac}[$try] = $schol;
              $SCHOLsched{$schol}[$try] = $fac;
              delete($FACslots{$fac}{$try});
              $nFACslots{$fac}--;
              $nFACsched{$fac}++;
              delete($SCHOLslots{$schol}{$try});
              $nSCHOLslots{$schol}--;
              $nSCHOLsched{$schol}++;
              $F2S{$fac}{$schol}++;
              $S2F{$schol}{$fac}++;
            }
          }
        }
      }
      foreach my $fac (sort {$nFACsched{$a} <=> $nFACsched{$b}} keys %nFACsched){
        if($nFACsched{$fac} < $minINTS){
          $minINTS = $nFACsched{$fac};
        }
        if($nFACsched{$fac} > $MINFacInt){
          last;
        }

        foreach my $schol (sort {$nSCHOLsched{$a} <=> $nSCHOLsched{$b}} keys %nSCHOLsched){
          if(exists($F2S{$fac}{$schol})){
            next;
          }
          if($nSCHOLsched{$schol} == $MAXInt){
            last;
          }
          for(my $try=0; $try<$MAXInt; $try++){
            if(exists($SCHOLslots{$schol}{$try}) && exists($FACslots{$fac}{$try})){
              ##print "\tFound a slot! fill it in!\n";
              $FACsched{$fac}[$try] = $schol;
              $SCHOLsched{$schol}[$try] = $fac;
              delete($FACslots{$fac}{$try});
              $nFACslots{$fac}--;
              $nFACsched{$fac}++;
              delete($SCHOLslots{$schol}{$try});
              $nSCHOLslots{$schol}--;
              $nSCHOLsched{$schol}++;
              $F2S{$fac}{$schol}++;
              $S2F{$schol}{$fac}++;
            }
          }
        }
      }

      foreach my $schol (sort {$nSCHOLsched{$a} <=> $nSCHOLsched{$b}} keys %nSCHOLsched){
        if($nSCHOLsched{$schol} < $minINTS){
          $minINTS = $nSCHOLsched{$schol};
        }
        else{
          last;
        }
      }
      foreach my $fac (sort {$nFACsched{$a} <=> $nFACsched{$b}} keys %nFACsched){
        if($nFACsched{$fac} < $minINTS){
          $minINTS = $nFACsched{$fac};
        }
        else{
          last;
        }
      }
      print "minINTS = $minINTS\n" if($VERBOSE);
    }
        
    ## check how many interviews everyone got
    @scholinthist = ();
    @facinthist = ();
    my @facfillhist = ();
    my %factoofew = ();
    $nfactoofew = 0;
    my %scholtoofew = ();
    $nscholtoofew = 0;
    for(my $i=0; $i<=$MAXInt; $i++){
      $scholinthist[$i] = 0;
      $facinthist[$i] = 0;
      $facfillhist[$i] = 0;
    }
    foreach my $fac (keys %Faculty){
      $facinthist[$nFACsched{$fac}]++;
      my $empty = $MAXInt-$nFACslots{$fac};
      $facfillhist[$empty]++;
      if($nFACsched{$fac} < $MINFacInt){
        $factoofew{$fac}++;
        $nfactoofew++;
      }
    }
    foreach my $schol (keys %Scholars){
      $scholinthist[$nSCHOLsched{$schol}]++;
      if($nSCHOLsched{$schol} < $MINScholInt){
        $scholtoofew{$schol}++;
        $nscholtoofew++;
      }
    }
    print "\n\n$nfactoofew faculty and $nscholtoofew scholars with too few interviews\n";
    print "Faculty with too few interviews:\n";
    foreach my $fac (sort keys %factoofew){
      print "\t$fac\n";
    }
    print "Scholars with too few interviews:\n";
    foreach my $schol (sort keys %scholtoofew){
      print "\t$schol\n";
    }
    print "\nHere is a table showing the number of faculty and scholars who have N interviews:\n";
    print "\nN\t#Faculty\t#Scholars\n";
    for(my $i=0; $i<=$MAXInt; $i++){
      print "$i\t$facfillhist[$i]\t$scholinthist[$i]\n";
    }
  }
}

print "made a schedule...\n";

## Let's print out the schedules!

my $totScholInts = 0;
my $totFacInts = 0;
my $GAPCHR = ",";
my $facSchedFile = "$DIROUT/CompleteFacultySchedule.csv";
open(OUT,">$facSchedFile") or die "cannot write to $facSchedFile\n";
print OUT "FacultyName";
for(my $i=0; $i<$MAXInt; $i++){
  print OUT "$GAPCHR$TIMES[$i]";
}
print OUT "\n";
foreach my $fac (sort {$facultyEMAILtoNAME{$a} cmp $facultyEMAILtoNAME{$b}} keys %Faculty){
  print OUT "$facultyEMAILtoNAME{$fac}";
  my $numInt = 0;
  for(my $i=0; $i<$MAXInt; $i++){
    my $schol = $FACsched{$fac}[$i];
    if($schol eq "NA"){
      print OUT "${GAPCHR}NA";
    }
    elsif($schol eq ""){
      print OUT "${GAPCHR}NA";
    }
    else{
      print OUT "${GAPCHR}$scholE2N{$schol}";
      $totFacInts++;
      $numInt++;
    }
  }
  print OUT "\n";
  if($nFACsched{$fac} < $MINFacInt){
    print "doh, $fac has $numInt interviews, but $nFACsched{$fac} sched and $nFACslots{$fac} slots\n";
    ##exit;
  }
}
print OUT "\n\n\n";
print OUT "Please find all zoom links here:\n";
print OUT sprintf '=HYPERLINK("%s")',
  "https://docs.google.com/spreadsheets/d/173_ZZX_MneFDMDl-e1KkIz1HNJ9T-g1BQuAySTBF4_A/edit#gid=0";
close(OUT);
print "There are a total of $totFacInts faculty interviews\n";

my $scholSchedFile = "$DIROUT/CompleteScholarSchedule.csv";
open(OUT,">$scholSchedFile") or die "cannot write to $scholSchedFile\n";
print OUT "ScholarName";
for(my $i=0; $i<$MAXInt; $i++){
  print OUT "${GAPCHR}$TIMES[$i]";
}
print OUT "\n";
##{$scholE2N{$a} cmp $scholE2N{$b}}
foreach my $schol (sort  keys %Scholars){
  print OUT "$scholE2N{$schol}";
  my $numInt = 0;
  for(my $i=0; $i<$MAXInt; $i++){
    if($SCHOLsched{$schol}[$i] ne ""){
      my $fac = $SCHOLsched{$schol}[$i];
      print OUT "${GAPCHR}";
      print OUT "$facultyEMAILtoNAME{$fac}";
      $totScholInts++;
      $numInt++;
    }
    else{
      print OUT "${GAPCHR}NA";
    }
  }
  print OUT "\n";
  if($numInt != $nSCHOLsched{$schol}){
    print "doh, $schol has $numInt interviews, but $nSCHOLsched{$schol} sched and $nSCHOLslots{$schol} slots\n";
    ##exit;
  }

}
print OUT "\n\n\n";
print OUT "Please find all zoom links here:\n";
print OUT sprintf '=HYPERLINK("%s")',
  "https://docs.google.com/spreadsheets/d/173_ZZX_MneFDMDl-e1KkIz1HNJ9T-g1BQuAySTBF4_A/edit#gid=0";
close(OUT);
print "There are a total of $totScholInts scholar interviews\n";

##print out individual scholar schedules
foreach my $schol (sort keys %Scholars){
  my $scholSchedFile = "$DIROUT/ScholarSchedules/$scholE2N{$schol}.csv";
  if(!(-d "$DIROUT/ScholarSchedules")){
    `mkdir $DIROUT/ScholarSchedules`;
  }

  open(OUT,">$scholSchedFile") or die "cannot write to $scholSchedFile\n";
  print OUT "ScholarName";
  for(my $i=0; $i<$MAXInt; $i++){
    print OUT "${GAPCHR}$TIMES[$i]";
  }
  print OUT "\n";
  
  print OUT "$scholE2N{$schol}";
  my $numInt = 0;
  for(my $i=0; $i<$MAXInt; $i++){
    if($SCHOLsched{$schol}[$i] ne ""){
      my $fac = $SCHOLsched{$schol}[$i];
      print OUT "${GAPCHR}";
      print OUT "$facultyEMAILtoNAME{$fac}";
      $numInt++;
    }
    else{
      print OUT "${GAPCHR}NA";
    }
  }
  print OUT "\n\n\n";
  print OUT "Please find all zoom links here:\n";
  print OUT sprintf '=HYPERLINK("%s")',
	      "https://docs.google.com/spreadsheets/d/173_ZZX_MneFDMDl-e1KkIz1HNJ9T-g1BQuAySTBF4_A/edit#gid=0";
  print OUT "\n";
  close(OUT);
}

#print out individual faculty schedules
foreach my $fac (sort keys %Faculty){
  my $facSchedFile = "$DIROUT/FacultySchedules/$facultyEMAILtoNAME{$fac}.csv";
  if(!(-d "$DIROUT/FacultySchedules")){
    `mkdir $DIROUT/FacultySchedules`;
  }

  open(OUT,">$facSchedFile") or die "cannot write to $facSchedFile\n";
  print OUT "FacultyName";
  for(my $i=0; $i<$MAXInt; $i++){
    print OUT "${GAPCHR}$TIMES[$i]";
  }
  print OUT "\n";
  print OUT "$facultyEMAILtoNAME{$fac}";
  my $numInt = 0;
  for(my $i=0; $i<$MAXInt; $i++){
    my $schol = $FACsched{$fac}[$i];
    if($schol eq "NA"){
      print OUT "${GAPCHR}NA";
    }
    elsif($schol eq ""){
      print OUT "${GAPCHR}NA";
    }
    else{
      print OUT "${GAPCHR}$scholE2N{$schol}";
      $numInt++;
    }
  }
  print OUT "\n\n\n";
  print OUT "Please find all zoom links here:\n";
  print OUT sprintf '=HYPERLINK("%s")',
	      "https://docs.google.com/spreadsheets/d/173_ZZX_MneFDMDl-e1KkIz1HNJ9T-g1BQuAySTBF4_A/edit#gid=0";
  print OUT "\n";
  close(OUT);
}

#print final zoom link list
if(1){
  my $faczoom = "$DIROUT/facZoomLinks.csv";
  open(OUT,">$faczoom") or die "cannot write to $faczoom\n";
  foreach my $fac (sort keys %Faculty){
    print OUT "$facultyEMAILtoNAME{$fac},$fac,$Faculty{$fac}[1]\n";
  }
  close(OUT);
}

exit;

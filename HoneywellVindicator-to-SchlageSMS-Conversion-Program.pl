#!C:\Perl\bin\perl.exe
################################################################################
# Honeywell Vindicator to Schlage SMS Program                                  #
# Copyright (c) 2013, Jason Unovitch (jason.unovitch@gmail.com and @us.af.mil) #
# Available at https://github.com/junovitch/HoneywellVindicator-to-SchlageSMS  #
#                                                                              #
# Redistribution and use in source and binary forms, with or without           #
# modification, are permitted provided that the following conditions are met:  #
#                                                                              #
#    (1) Redistributions of source code must retain the above copyright        # 
#    notice, this list of conditions and the following disclaimer.             #
#                                                                              #
#    (2) Redistributions in binary form must reproduce the above copyright     #
#    notice, this list of conditions and the following disclaimer in the       #
#    documentation and/or other materials provided with the distribution.      #
#                                                                              #
# THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR IMPLIED #
# WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF         #
# MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO   #
# EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,       #
# SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, #
# PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;  #
# OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,     #
# WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR      #
# OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF       #
# ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.                                   #
################################################################################
# Usage:
# See under $instructions below.
#
# Revision History:
# V1.00	JU - 20120819 - Designed and tested on Linux and tested on Windows XP
#	with ActiveState Perl 5.16 via VirtualBox.
# V1.01	JU - 20120821 - Updated with some error control to deal with invalid
#	fields.
# V1.02	JU - 20120921 - Updated with significantly better error control, cleaner
#	math, and descriptive comments.
# V1.03	JU - 20121022 - Defined a next to ignore badges that aren't valid, a 
#	catch to print the headings properly, and some flow control for badges
#	that use facilty code 22.
# V1.10	JU - 20121023 - Increment revision number due to interactive GUI.
#	If run with a valid CLI arg it works like normal.  If no CLI arg is
#	given or it is double clicked, it runs a GUI.
# V1.11	JU - 20130202 - Cleaned up all code significantly, worked out kinks and
#	got use strict finally turned back on. Working on Access db integration.
# V1.20	JU - 20130219 - Increment revesion number because of new "diff" function
#	that allows showing only the changes since the prior run.
# V1.21	JU - 20130312 - Major code refactoring/cleanup
# V1.22	JU - 20130318 - Major code refactoring/cleanup
# V1.30	JU - 20130320 - Changed to import Excel formats and remove intermediary
# 	steps to open with Excel and "save as" a .csv file. The order of the
# 	Excel no longer matters and if a column, say Middle Initial, is removed
#	then the program will just print empty spaces and will not error out. 
# V1.31	JU - 20130401 - Cleanup Tkx GUI grid code and prep for a status display
#	inline with the GUI.
# V1.32 JU - 20130402 - Migrate GUI grid to all object oriented and finalize
#	status display in GUI.
# V1.33 JU - 20130403 - Documentation updates and bug fixes.
# V1.34 JU - 20130404 - Updated to use Personnel Group and Authority Level to
#	better determine who can access which doors.
# V1.35 JU - 20130408 - Reorganized categories slightly for clarity.
# V1.36 JU - 20130419 - Minor bug fixes and cleanup of commented draft code.
# V1.40 JU - 20130421 - Updated to use hash of hash structure instead of a flat
# 	file. Makes determining what has changed easier to manage.
# V1.41 JU - 20130428 - Streamlined hash of hash comparison code as well as
#	updated to display last, first for each changed record.
# V1.42 JU - 20130528 - Remove Middle Initial field and publish to a permanent
#	repository.
#
################################################################################
##  Perl Module Declaration  ###################################################
################################################################################

use v5.14;
use warnings;
use strict;
use Tkx;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Getopt::Std;
use Data::Dumper;

################################################################################
##  Variable Declaration  ######################################################
################################################################################

# Define script version number
our $VERSION = "1.42";
(my $progname = $0) =~ s,.*[\\/],,;

# Setup Date-stamps
my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst)=localtime(time);
my $datestamp = sprintf("%4d%02d%02d%02d%02d%02d", $year+1900,$mon+1,$mday,$hour,$min,$sec);

# Define default save locations and names of temporary files used
my $target_directory = "$ENV{'USERPROFILE'}\\Desktop";

# Get list of files in import directory in case older changes haven't been made.
opendir DIR, $target_directory or die "cannot open dir $target_directory: $!";
my @target_directory_files = readdir DIR;
closedir DIR;

# Set the default names of files used
my $baseline_import_file = "$target_directory\\BASELINE-SMS-import.txt";
my $baseline_raw_datafile = "$target_directory\\BASELINE-SMS-rawdata.txt";
my $output_full_import_file = "$target_directory\\${datestamp}-SMS-fullimport.txt";
my $output_raw_datafile = "$target_directory\\${datestamp}-SMS-rawdata.txt";
my $output_differential_import_file = "$target_directory\\${datestamp}-SMS-import-this-file.txt";
my $manual_update_filename = "$target_directory\\${datestamp}-SMS-manual-updates.txt";

# Declare variables used across the namespace of several different
# subroutines
my @cleanup_list;
my $log;
my $input_filename;
my $has_gui = "NO";

# Instructions to be displayed when help is requested.
my $instructions = <<"__END__";
Usage:
Run with no arguements for the GUI.
Run $progname -f importfile.xls to convert via the command line.
Run $progname -v to see the version number.
Run $progname -h to see this help.

To prepare the import file, do the following 
Export Vindicator Database:
  1. Click Report Wizard
    a. Next
    b. Select Badge Report (Default)
    c. Next
    d. Select all fields and move to report fields
    e. Finish
    f. Select Export Excel
Convert Vindicator Database for Import
  1. Open $progname on desktop
    a. Select Open XLS File and browse to the Excel file.
    b. Click Perform Conversion
    d. Click Save and Exit
Upkeep of Standalone SMS Database
  1. Perform Export Vindicator Database steps
  2. Perform Convert Vindicator Database steps follow the on screen instructions.
  3. Import into SMS
    a. Select Facilities => Import => User Input from Text File
    b. Browse to $baseline_import_file
    c. Select "Perl Import" under layout templates
    d. Select "Import users"
__END__

################################################################################
##  General Program Logic  #####################################################
################################################################################

my %opt;
getopts("f:hv", \%opt);
if (%opt) {
	if ($opt{f}) {
		$input_filename = $opt{f};
		&main;
		&cleanup;
	} elsif ($opt{h}) {
		say "$instructions";
		exit;
	} elsif ($opt{v}) {
		say "$progname v$VERSION";
		exit;
	} else {
		exit;
	}
} else {
	&start_gui;
}

################################################################################
##  GUI Shell Code  ############################################################
################################################################################
# Ref:  http://docs.activestate.com/activeperl/5.16/lib/Tkx/Tutorial.html
# 2005 Activestate
################################################################################

sub start_gui {
	$has_gui = "YES";
	
	Tkx::package_require("style");
	Tkx::style__use("as", -priority => 70);
	
	our $mw = Tkx::widget->new(".");
	$mw->configure(-menu => mk_menu($mw));
	Tkx::tk(appname => "V5 to Schlage Management Software Conversion Program");
	
	sub mk_menu {
		my $mw = shift;
		my $menu = $mw->new_menu;
		my $file = $menu->new_menu(
			-tearoff => 0,
		);
		$menu->add_cascade(
			-label => "File",
			-underline => 0,
			-menu => $file,
		);
		$file->add_command(
			-label => "Open XLS",
			-underline => 0,
			-accelerator => "Ctrl+O",
			-command => \&get_file,
		);
		$mw->g_bind("<Control-o>", \&get_file);
		$file->add_command(
			-label => "Perform Conversion",
			-underline => 0,
			-accelerator => "Ctrl+P",
			-command => \&main,
		);
		$mw->g_bind("<Control-p>", \&main);
		$file->add_command(
			-label   => "Save and Exit",
			-underline => 1,
			-command => sub {cleanup();},
		);	
		$file->add_command(
			-label   => "Exit Without Saving",
			-underline => 1,
			-command => [\&Tkx::destroy, $mw],
		);
		my $help = $menu->new_menu(
			-name => "help",
			-tearoff => 0,
		);
		$menu->add_cascade(
	 		-label => "Help",
			-underline => 0,
			-menu => $help,
		);
		$help->add_command(
			-label => "\u$progname Manual",
			-command => [\&print, $instructions],
		);
		
		my $about_menu = $help;
	
		$about_menu->add_command(
			-label => "About \u$progname",
			-command => \&about,
		);
	
		return $menu;
	}
	
	sub about {
		Tkx::tk___messageBox(
			-parent => $mw,
			-title => "About \u$progname",
			-type => "ok",
			-icon => "info",
			-message => "$progname v$VERSION\n" .
				"Copyright 2012-2013 Jason Unovitch. " .
				"All rights reserved.",
			);
	}
	
################################################################################
##  Site-local GUI Code  #######################################################
################################################################################
# Ref:  http://www.tkdocs.com/tutorial/index.html
################################################################################
	
	# First we are going to define all of our GUI "stuff" and outline what
	# we want each label or button to say.	
	$mw->g_wm_protocol('WM_DELETE_WINDOW' => \&confirm_exit);
	my $content = $mw->new_ttk__frame(-padding => "5 5 5 5");

	my $srcFilePrefix = $content->new_ttk__label(
		-text => "Source Filename: ",
	);
	my $srcFileName = $content->new_ttk__label(
	-textvariable => \$input_filename,
	);
	my $pickInputButton = $content->new_ttk__button(
		-text => "1.  Open XLS",
		-width => 30,
		-command => sub {get_file();},
	);
	my $outFilePrefix = $content->new_ttk__label(
		-text => "Output Filename: ",
	);
	my $outFileName = $content->new_ttk__label(
		-textvariable => \$output_full_import_file,
	);
	my $doConvertButton = $content->new_ttk__button(
		-text => "2.  Perform Conversion",
		-width => 30,
		-command => sub {main();},
	);
	my $quitButton = $content->new_ttk__button(
		-text => "3.  Save and Exit",
		-width => 30,
		-command => sub {cleanup();},
	);
	$log = $content->new_tk__text(
		-state => "disabled",
		-width => "120",
		-height => "36",
		-wrap => "none",
	);
	my $logScrollBar = $content->new_ttk__scrollbar(
		-command => [$log, "yview"],
		-orient => "vertical",
	);
	$log->configure(-yscrollcommand => [$logScrollBar, 'set']);

	# Now we are going to take all that GUI "stuff" and toss it into a grid.
	$content->g_grid	(-row => 0, -column => 0, -sticky => "nwes");
	$srcFilePrefix->g_grid	(-row => 0, -column => 0, -sticky => "w");
	$srcFileName->g_grid	(-row => 0, -column => 1, -sticky => "w", -padx => "15");
	$pickInputButton->g_grid(-row => 0, -column => 2, -sticky => "e");
	$outFilePrefix->g_grid	(-row => 1, -column => 0, -sticky => "w");
	$outFileName->g_grid	(-row => 1, -column => 1, -sticky => "w", -padx => "15");
	$doConvertButton->g_grid(-row => 1, -column => 2, -sticky => "e");
	$quitButton->g_grid	(-row => 2, -column => 2, -sticky => "e");
	$log->g_grid		(-row => 3, -columnspan => 3, -sticky => "we");
	$logScrollBar->g_grid	(-row => 3, -column => 4, -sticky => "ns");

	sub confirm_exit {
		my $answer = Tkx::tk___messageBox(
			-parent => $mw,
			-title => "Warning",
			-type => "yesno",
			-icon => "info",
			-message => "Are you sure you want to exit?\n" .
				"If you have run the conversion, please use the save and exit button.",
			);
		if ($answer eq 'yes') {
			exit;
		}
	}

	sub get_file {
		$input_filename = Tkx::tk___getOpenFile();
	}

	sub wrong_file {
		Tkx::tk___messageBox(
			-parent => $mw,
			-title => "Warning",
			-type => "ok",
			-icon => "info",
			-message => "Please select an *.xls file",
		);
	}	

	Tkx::MainLoop();
	exit;
}

################################################################################
##  Backend Subroutine  ########################################################
################################################################################

sub print {
	# Decide if a GUI is in use or not then prints as appropriate
	if ($has_gui =~ /YES/) {
		my ($_) = @_;
		my $numlines = $log->index("end - 1 line");
		$log->configure(-state => "normal");
		if ($log->index("end-1c")!="1.0") {$log->insert_end("\n");}
		$log->insert_end($_);
		$log->configure(-state => "disabled");
		Tkx::update();
		say("@_");
	} else {
		say("@_");
	}
}

sub main {
	# Check to see that only an XLS file is opened.
	unless ($input_filename =~ /\.xls$/i) {
		&print("Needs to be an Excel .xls file!");
		&wrong_file && return if ($has_gui =~ /YES/); 
		exit if ($has_gui =~ /NO/);
	}

	# Display up front if there are left over changes from a prior run that
	# didn't save and exit properly.
	foreach my $file (@target_directory_files) {
		if ($file =~ /manual-updates.txt/) {
			&print("WARNING: There appears to be unmerged updates from a previous run in $file.");
			open(FILE1, "<", "$target_directory\\$file") || die $?;
			while(<FILE1>) {
				chomp $_; &print("$_");
			}
			close(FILE1);
		}
	}

	# Access Excel via the OLE interface and open $input_filename
	$Win32::OLE::Warn = 3;
	my $Excel = Win32::OLE->new('Excel.Application');
	my $workbook = $Excel->Workbooks->Open($input_filename);
	my $Sheet = $workbook->Worksheets(1);
	$Sheet->Activate();

	# Check the number of columns so that program logic knows stopping points
	my $lastColumn = $Sheet->UsedRange->Find({What => "*", 
		SearchDirection	=> xlPrevious,
		SearchOrder	=> xlByColumns})->{Column};
	&print("==>> Total Columns to process is $lastColumn...");

	# Check the number or rows so that program logic knows stopping points
	my $lastRow = $Sheet->UsedRange->Find({What => "*",
		SearchDirection	=> xlPrevious,
		SearchOrder	=> xlByRows})->{Row};
	&print("==>> Total Rows to process is $lastRow...");

	# Define variables to cover each column needed.
	my ($Last_Name,
		$First_Name,
		$Issue_Date,
		$Expire_Date,
		$V5_Personnel_Group,
		$V5_Authority_Level,
		$V5_Privilege,		
		$V5_Card_Active,
		$V5_Normal_PIN,
		$V5_Card_Number,
	);

	# Parse each column and load it into the appropriate variable based off the header
	foreach my $column (1..$lastColumn) {
		my $header = $Sheet->Cells(1,$column)->{'Value'};
		given ($header) {
			when (/^Last Name$/i) {
				$Last_Name = $Sheet->Columns($column)->{'Value'};
			}
			when (/^First Name$/i) {
				$First_Name = $Sheet->Columns($column)->{'Value'};
			}
			when (/^Issue Date$/i) {
				$Issue_Date = $Sheet->Columns($column)->{'Value'};
			}
			when (/^Expire Date$/i) {
				$Expire_Date = $Sheet->Columns($column)->{'Value'};
			}
			when (/^V5 Personnel Group$/i) {
				$V5_Personnel_Group = $Sheet->Columns($column)->{'Value'};
			}
			when (/^V5 Authority Level$/i) {
				$V5_Authority_Level = $Sheet->Columns($column)->{'Value'};
			}
			when (/^V5 Privilege$/i) {
				$V5_Privilege = $Sheet->Columns($column)->{'Value'};
			}
			when (/^V5 Card Active$/i) {
				$V5_Card_Active = $Sheet->Columns($column)->{'Value'};
			}
			when (/^V5 Normal PIN$/i) {
				$V5_Normal_PIN = $Sheet->Columns($column)->{'Value'};
			}			
			when (/^V5 Card Number$/i) {
				$V5_Card_Number = $Sheet->Columns($column)->{'Value'};
			}
		}
	}
	
	# Create some Perl arrays to hold the actual data for program use.
	my (@Last_Names,
		@First_Names,
		@Issue_Dates,
		@Expire_Dates,
		@V5_Personnel_Groups,
		@V5_Authority_Levels,
		@V5_Privileges,
		@V5_Card_Actives,
		@V5_Normal_PINs,
		@V5_Card_Numbers,
		@SMS_Card_Numbers,
	);

	# Parse through each $string reference to load data into a Perl usable form
	# in an array that is named very similarly.  Realize that the $First_Name ...
	# titled variables are competely separate from the @First_Names variables and
	# they are only named similarly for us to follow data flow through the program.
	foreach my $lastNameRefArray (@$Last_Name) {
		foreach (@$lastNameRefArray) {
			if (defined ($_)) { push(@Last_Names, $_); } else { push(@Last_Names, ''); }
		}
	}

	foreach my $firstNameRefArray (@$First_Name) {
		foreach (@$firstNameRefArray) {
			if (defined ($_)) { push(@First_Names, $_); } else { push(@First_Names, ''); }
		}
	}

	foreach my $issueDatesRefArray (@$Issue_Date) {
		foreach (@$issueDatesRefArray) {
			if (defined ($_)) { push(@Issue_Dates, $_); } else { push(@Issue_Dates, ''); }
		}
	}

	foreach my $expireDatesRefArray (@$Expire_Date) {
		foreach (@$expireDatesRefArray) {
			if (defined ($_)) { push(@Expire_Dates, $_); } else { push(@Expire_Dates, ''); }
		}
	}	
	
	foreach my $V5PersonnelGroupRefArray (@$V5_Personnel_Group) {
		foreach (@$V5PersonnelGroupRefArray) {
			if (defined ($_)) { push(@V5_Personnel_Groups, $_); } else { push(@V5_Personnel_Groups, ''); }
		}
	}

	foreach my $V5AuthorityLevelRefArray (@$V5_Authority_Level) {
		foreach (@$V5AuthorityLevelRefArray) {
			if (defined ($_)) { push(@V5_Authority_Levels, $_); } else { push(@V5_Authority_Levels, ''); }
		}
	}

	foreach my $V5PrivilegesRefArray (@$V5_Privilege) {
		foreach (@$V5PrivilegesRefArray) {
			if (defined ($_)) { push(@V5_Privileges, $_); } else { push(@V5_Privileges, ''); }
		}
	}

	foreach my $V5CardActiveRefArray (@$V5_Card_Active) {
		foreach (@$V5CardActiveRefArray) {
			if (defined ($_)) { push(@V5_Card_Actives, $_); } else { push(@V5_Card_Actives, ''); }
		}
	}	
	
	foreach my $V5NormalPINsRefArray (@$V5_Normal_PIN) {
		foreach (@$V5NormalPINsRefArray) {
			if (defined ($_)) { push(@V5_Normal_PINs, $_); } else { push(@V5_Normal_PINs, ''); }
		}
	}
	
	# Set some unique variables used in the calculation of SMS style badge numbers
	my ($SMS_badge_code,
		$binary_V5_badge_code,
		$facility_code_and_binary_V5_badge_code,
		$first12bits,
		$last12bits,
		$even_parity_check,
		$leading_parity_bit,
		$odd_parity_check,
		$trailing_parity_bit,
	);

	# In this section, we are going to process our badge numbers out of Honeywell's
	# Vindicator program and convert it into the format Schlage Management Software
	# understands. This was all built by reverse engineering information off the
	# badges and we'll walk it through one step at a time.
	# 
	# This whitepaper was very helpful in understanding the formats used by prox badges
	# http://www.isecuretech.com/download/SmartCardReader/OMNIKEY/driver/OK5x21/OK5x25_Prox_ATRDecode.pdf
	#
	# We started with a hex dump from a badge, 3B 06 01 02 55 03 13 17
	# We knew that the Vindicator number and number on the badge was 31317
	#
	# At this point, based off the spec sheet, our 31317 number is the "Card Number".
	# The 0255 just before is the "Facility Code", which Honeywell's Vindicator does
	# not use. This was discovered the hard way based off a hex dump of badges over
	# 62199.  Those badges use a different facility code and did not work initially.
	# As such there is a section that joins the appropriate facility number.
	#
	# At this point, we have 24 information bits of a 26 bit format. The bits before
	# and after are for calculating parity.
	foreach my $decimalBadgeNumberRefArray (@$V5_Card_Number) {
		foreach my $badgeNumber (@$decimalBadgeNumberRefArray) {
			if (defined($badgeNumber) && ($badgeNumber =~ /^\d+$/ && $badgeNumber > 0 && $badgeNumber < 65535)) {
				# It's defined,       it's a number,             it's greater than 0 and less than 65535
				#
				# Example Input Data: "31317".
				# Convert "n", a 16 bit big endian value from decimal
				# To "B16", a bit string in descending order
				$binary_V5_badge_code = unpack("B16", pack ("n", $badgeNumber));
				# Example Output: 0111101001010101 
				#
				# Join facility code (22 decimal, 00010110 binary, for badges above 62199.
				# Join 255 decimal, 11111111 binary as default for all other  badges
				if ($badgeNumber > 62199) {
					$facility_code_and_binary_V5_badge_code = join("", "00010110", $binary_V5_badge_code);
				} else {
					$facility_code_and_binary_V5_badge_code = join("", "11111111", $binary_V5_badge_code);
				}
				# Example Output: 111111110111101001010101
				#
				# Split in 12 bit halves for parity check
				($first12bits, $last12bits) = unpack("(A12)*(A12)", $facility_code_and_binary_V5_badge_code);
				# Example Output: 111111110111 and 101001010101
				#
				# Translate spaces to nothing, automatically converts back to decimal
				# Essentially this counts the number of 1's
				$even_parity_check = ($first12bits =~ tr/1//);
				$odd_parity_check = ($last12bits =~ tr/1//);
				# Example Output: 11 for $even_parity_check and 6 for $odd_parity_check				
				#
				# Set default of 0, reset to one when $even_parity_check is odd.
				# This divides by 2 and if there is a remainder it will set the
				# parity to 1.
				$leading_parity_bit = 0; $leading_parity_bit = 1 if($even_parity_check % 2);
				# Example Output: 1  (11/2's remainder is 1, since this is valid it resets)
				#
				# Set default of 0, reset to one when $odd_parity_check is even.
				# This divides by 2 and if there is no remainder it will set the
				# parity to 1.
				$trailing_parity_bit = 0; $trailing_parity_bit = 1 if ! ($odd_parity_check % 2);						
				# Example Output: 1 (6/2's remainder is null, so the 'if !' inverses the usual meaning)
				#
				# Octal Conversions
				# Joins each field to do a string print as a 16 digit octal number
				# for Schlage Management Software to use. Finally joins blank
				# spaces into 0's just to prevent issues during import.
				$SMS_badge_code = sprintf('%16O', oct(join("", "0b1",
					$leading_parity_bit,
					$facility_code_and_binary_V5_badge_code,
					$trailing_parity_bit)));
				$SMS_badge_code =~ tr/ /0/;
				# Example Output: 0000000777572253
				# Push the end result back into the final array for processing
				push(@SMS_Card_Numbers, $SMS_badge_code);
				push(@V5_Card_Numbers, $badgeNumber);

			} elsif (defined($badgeNumber) && ($badgeNumber =~/V5 Card Number/)) {
				# Prints header for the first line
				push(@SMS_Card_Numbers, "SMS Badge Code");
				push(@V5_Card_Numbers, "V5 Card Number");
			} else {
				# Prints 0s for anything not valid
				push(@SMS_Card_Numbers, "0");
				push(@V5_Card_Numbers, "0");
			}
		}
	}

	# Generate a hash of hash structure to hold the badge information
	# Note, do not re-order this or remove the alphabetical prefix as
	# it will effect the order of the import template used in SMS. 
	my %badgeHash;
	for my $element ( 0 .. $lastRow ) {
		if (($V5_Privileges[$element] =~ /NORMAL USER/i) && ! ($V5_Authority_Levels[$element] =~ /^$/)) {
			$badgeHash{ $V5_Card_Numbers[$element] } = {
				"A-Last Name"		=> $Last_Names[$element],
				"B-First Name"		=> $First_Names[$element],
				"C-Issue Date"		=> $Issue_Dates[$element]->Win32::OLE::Variant::Date("MM/dd/yyyy"),
				"D-Expire Date"		=> $Expire_Dates[$element]->Win32::OLE::Variant::Date("MM/dd/yyyy"),
				"E-Auth Group"		=> $V5_Personnel_Groups[$element] . " " .  $V5_Authority_Levels[$element],
				"F-V5 Card Active"	=> $V5_Card_Actives[$element],
				"G-V5 Normal PIN"	=> $V5_Normal_PINs[$element],
				"H-SMS Card Number"	=> $SMS_Card_Numbers[$element],
			};
		}
	}

	# Here we'll dump a raw copy of our data for future comparisons
	$Data::Dumper::Purity = 1;
	$Data::Dumper::Sortkeys = 1;
	open(NEWRAWFH, ">", "$output_raw_datafile") || die $?;
	print NEWRAWFH Data::Dumper->Dump([\%badgeHash], ['*oldbadgeHash']);
	close(NEWRAWFH);
	&print("Processed " . scalar(keys %badgeHash) .  " NORMAL USER records with non-blank permissions...");
	
	# With all the data on hand, we'll open up our dated output
	# filename, sort through our badge off the %badgeHash key 
	# (badge number), then print each value into the file in a
	# CSV separated format.
	my $active_count;
	open(CSVFH, ">", "$output_full_import_file") || die $?;
	foreach my $badge ( sort keys %badgeHash ) {
		if (($badgeHash{$badge}{'F-V5 Card Active'} eq 1)) {
			print CSVFH "$badge";
			foreach ( sort keys %{ $badgeHash{$badge} } ) {
				print CSVFH ",$badgeHash{$badge}{$_}";
			}
			print CSVFH "\n";
			$active_count++;
		}
	}
	close(CSVFH);
	&print("Saved " . $active_count .  " V5 CARD ACTIVE records...");

	# Don't save Excel session, just exit
	$Excel->{DisplayAlerts} = 0;
	$workbook->{Saved} = 0;
	$Excel->Quit;
	&print("DONE");

	# If there is a baseline file, display the differences to the user
	# for update into the main database. All temporary files will be
	# added to the cleanup list and we'll wait for the user to confirm
	# all updates have been made.
	if (-e $baseline_raw_datafile && -e $output_raw_datafile) {
		&print("==>> Accessing old badge records...");
		open(OLDRAWFH, "<", $baseline_raw_datafile) || die $?;
		my %oldbadgeHash;
		{
			local $/;
			%oldbadgeHash = eval <OLDRAWFH>;
		}
		close(OLDRAWFH);
		&print("DONE");
	
		&print("==>> Checking for what has changed...");
		open(MANUALDIFFS, ">", "$manual_update_filename") || die $?;
		open(IMPORTDIFFS, ">", "$output_differential_import_file") || die $?;

		for my $badge ( keys %badgeHash) {
			if (exists $oldbadgeHash{$badge}) {
				
				# Say deactivated and move on to the next record if the record was
				# changed from active to in-active.
				if (($badgeHash{$badge}{'F-V5 Card Active'} eq 0) && ($oldbadgeHash{$badge}{'F-V5 Card Active'} eq 1)) {
					say MANUALDIFFS
					"$oldbadgeHash{$badge}{'A-Last Name'}, " .
					"$oldbadgeHash{$badge}{'B-First Name'} " .
					"(badge number $badge) deactivated. Delete from SMS.";
					next;
				}

				# Say needs manual review if the badge was in-active and went back
				# to active.
				if (($badgeHash{$badge}{'F-V5 Card Active'} eq 1) && ($oldbadgeHash{$badge}{'F-V5 Card Active'} eq 0)) {
					say MANUALDIFFS 
					"$oldbadgeHash{$badge}{'A-Last Name'}, " .
					"$oldbadgeHash{$badge}{'B-First Name'} " .
					"(badge number $badge) reactivated? Needs manual review.";
					next;
				}

				# Ignore any changes if the badge stays inactive. This may otherwise
				# display permission changes or date changes that have been done to
				# streamline or cleanup inactive records.
				if (($badgeHash{$badge}{'F-V5 Card Active'} eq 0) && ($oldbadgeHash{$badge}{'F-V5 Card Active'} eq 0)) {
					next;
				}

				# Finally, go through each change for badges that were and still are
				# active records.
				for my $key ( sort keys $badgeHash{$badge} ) {
					unless ($badgeHash{$badge}{$key} eq $oldbadgeHash{$badge}{$key}) {
						say MANUALDIFFS 
						"$oldbadgeHash{$badge}{'A-Last Name'}, " .
						"$oldbadgeHash{$badge}{'B-First Name'} " .
						"(badge number $badge) $key changed from " .
						"$oldbadgeHash{$badge}{$key} " .
						"to $badgeHash{$badge}{$key}";
					}
				}

			} else {

				# This will be hit if the record did not exist at all in the past run.
				# It will create an entry in CSV form within a text file for direct
				# import into SMS>
				print IMPORTDIFFS "$badge";
				foreach ( sort keys %{ $badgeHash{$badge} } ) {
					print IMPORTDIFFS ",$badgeHash{$badge}{$_}";
				}
				print IMPORTDIFFS "\n";
			}
		}

		# Close each file that was opened for writing.
		close(MANUALDIFFS);
		close(IMPORTDIFFS);
		&print("DONE");

		# Re-open the file for reading to say what changes are needed.
		&print("==>> Update the following changes by hand (if any)...");
		open(MANUALDIFFS, "<", "$manual_update_filename") || die $?;
		while(<MANUALDIFFS>) {
			chomp $_; &print("$_");
		}
		close(MANUALDIFFS);

		# Re-open the file for reading to say what changes are needed.
		&print("==>> Import $output_differential_import_file for the following new badges (if any)");
		open(IMPORTDIFFS, "<", "$output_differential_import_file") || die $?;
		while(<IMPORTDIFFS>) {
			chomp $_; &print("$_");
		}
		close(IMPORTDIFFS);

		# Prompt for user validation if we happen to be using the CLI
		if ($has_gui =~ /NO/) {
			&print("==>> Please type \"Yes\" to cleanup once changes have been made.");
			while (<STDIN>) {
				last if ($_ =~ /yes|y/i);
				&print("==>> Please enter yes once changes have been made.");
			}	
		}

		# Mark the files used for cleanup and say all done.
		push(@cleanup_list, "$manual_update_filename");
		push(@cleanup_list, "$output_differential_import_file");
		&print("");
		&print("ALL DONE!  Please save and exit once any updates above have been made.");

	} else {
		# If there is no baseline file, we'll rename our temporary file as the baseline file
		&print("==>> No Baseline file found!");
		&print("");
		&print("ALL DONE!  Please save and exit.");
	}

	sub cleanup {
		# First we'll rename the replacement file as the new baseline
		&print("==>> Renaming $output_full_import_file to $baseline_import_file...");
		rename($output_full_import_file, $baseline_import_file)
			&& &print("DONE")
			|| &print("unable to rename");
		&print("==>> Renaming $output_raw_datafile to $baseline_raw_datafile...");
		rename($output_raw_datafile, $baseline_raw_datafile)
			&& &print("DONE")
			|| &print("unable to rename");

		# Make a list of files to cleanup
		# Only applied for files already in the directory from before.
		foreach my $file (@target_directory_files) {
			if ($file =~ /\d{14}-SMS-(fullimport|rawdata|import-this-file|manual-updates)\.txt/) {
				push(@cleanup_list, "$target_directory\\$file");
			}
       		}

		# Then we'll remove each file listed on the cleanup list.
		&print("==>> Removing temporary files...");
		foreach(@cleanup_list) {
			unlink("$_")
				&& &print("Removed $_")
				|| &print("unable to remove $_");
		}
		&print("DONE");
		exit 0;
	}
}

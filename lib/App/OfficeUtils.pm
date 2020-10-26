package App::OfficeUtils;

# AUTHORITY
# DATE
# DIST
# VERSION

use 5.010001;
use strict;
use warnings;
use Log::ger;

our %SPEC;

our %args_libreoffice = (
    libreoffice_path => {
        schema => 'filename*',
        tags => ['category:libreoffice'],
    },
);

our %arg0_input_file = (
    input_file => {
        summary => 'Path to input file',
        schema => 'filename*',
        req => 1,
        pos => 0,
    },
);

our %arg1_output_file = (
    output_file => {
        summary => 'Path to output file',
        schema => 'filename*',
        pos => 1,
        description => <<'_',

If not specified, will output to stdout.

_
    },
);

our %arg1_output_file_or_dir = (
    output_file_or_dir => {
        summary => 'Path to output file or directory',
        schema => 'pathname*',
        req => 1,
        pos => 1,
        description => <<'_',

If not specified, will output to stdout.

_
    },
);

our %argopt_overwrite = (
    overwrite => {
        schema => 'bool*',
        cmdline_aliases => {O=>{}},
    },
);

$SPEC{officewp2txt} = {
    v => 1.1,
    summary => 'Convert Office word-processor format file (.doc, .docx, .odt, etc) to .txt',
    description => <<'_',

This utility uses one of the following backends:

* LibreOffice

_
    args => {
        %arg0_input_file,
        %arg1_output_file,
        %argopt_overwrite,
        %args_libreoffice,
        return_output_file => {
            summary => 'Return the path of output file instead',
            schema => 'bool*',
            description => <<'_',

This is useful when you do not specify an output file but do not want to show
the converted document to stdout, but instead want to get the path to a
temporary output file.

_
        },
        fmt => {
            summary => 'Run Unix fmt over the txt output',
            schema => 'bool*',
        },
    },
};
sub officewp2txt {
    my %args = @_;

    require File::Copy;
    require File::Temp;
    require File::Which;
    require IPC::System::Options;

  USE_LIBREOFFICE: {
        my $libreoffice_path = $args{libreoffice_path} //
            File::Which::which("libreoffice") //
              File::Which::which("soffice");
        unless (defined $libreoffice_path) {
            log_debug "libreoffice is not in PATH, skipped trying to use libreoffice";
            last;
        }

        my $input_file = $args{input_file};
        $input_file =~ /(.+)\.(\w+)\z/ or return [412, "Please supply input file with extension in its name (e.g. foo.doc instead of foo)"];
        my ($name, $ext) = ($1, $2);
        $ext =~ /\Ate?xt\z/i and return [304, "Input file '$input_file' is already text"];
        my $output_file = $args{output_file};

        if (defined $output_file && -e $output_file && !$args{overwrite}) {
            return [412, "Output file '$output_file' already exists, not overwriting (use --overwrite (-O) to overwrite)"];
        }

        my $tempdir = File::Temp::tempdir(CLEANUP => !$args{return_output_file});
        my ($temp_fh, $temp_file) = File::Temp::tempfile(undef, SUFFIX => ".$ext", DIR => $tempdir);
        (my $temp_out_file = $temp_file) =~ s/\.\w+\z/.txt/;
        File::Copy::copy($input_file, $temp_file) or do {
            return [500, "Can't copy '$input_file' to '$temp_file': $!"];
        };
        # XXX check that $temp_file/.doc/.txt doesn't exist yet
        IPC::System::Options::system(
            {die=>1, log=>1},
            $libreoffice_path, "--headless", "--convert-to", "txt:Text (encoded):UTF8", $temp_file, "--outdir", $tempdir);

      FMT: {
            last unless $args{fmt};
            return [412, "fmt is not in PATH"] unless File::Which::which("fmt");
            my $stdout;
            IPC::System::Options::system(
                {die=>1, log=>1, capture_stdout=>\$stdout},
                "fmt", $temp_out_file,
            );
            open my $fh, ">" , "$temp_out_file.fmt" or return [500, "Can't open '$temp_out_file.fmt': $!"];
            print $fh $stdout;
            close $fh;
            $temp_out_file .= ".fmt";
        }

        if (defined $output_file || $args{return_output_file}) {
            if (defined $output_file) {
                File::Copy::copy($temp_out_file, $output_file) or do {
                    return [500, "Can't copy '$temp_out_file' to '$output_file': $!"];
                };
            } else {
                $output_file = $temp_out_file;
            }
            return [200, "OK", $args{return_output_file} ? $output_file : undef];
        } else {
            open my $fh, "<", $temp_out_file or return [500, "Can't open '$temp_out_file': $!"];
            local $/;
            my $content = <$fh>;
            close $fh;
            return [200, "OK", $content, {"cmdline.skip_format"=>1}];
        }
    }

    [412, "No backend available"];
}

$SPEC{officess2csv} = {
    v => 1.1,
    summary => 'Convert Office spreadsheet format file (.ods, .xls, .xlsx) to one or more CSV files',
    description => <<'_',

This utility uses <pm:Spreadsheet::XLSX> to extract cell values of worksheets
and put them in one or more CSV file(s). If spreadsheet format is not .xlsx
(e.g. .ods or .xls), it will be converted to .xlsx first using Libreoffice
(headless mode).

You can select one or more worksheets to export. If unspecified, the default is
the first worksheet only. If you specify more than one worksheets, you need to
specify output *directory* instead of *output* file.

_
    args => {
        %arg0_input_file,
        %arg1_output_file_or_dir,
        %argopt_overwrite,
        %args_libreoffice,
        # XXX option to merge all csvs as a single file?
        worksheets => {
            summary => 'Select which worksheet(s) to convert',
            'x.name.is_plural' => 1,
            'x.name.singular' => 'worksheet',
            schema => ['array*', of=>'str*'],
            cmdline_aliases => {s=>{}},
        },
        all_worksheets => {
            summary => 'Convert all worksheets in the workbook',
            schema => 'true*',
            cmdline_aliases => {a=>{}},
        },
        always_dir => {
            summary => 'Assume output_file_or_dir is a directory even though there is only one worksheet',
            schema => 'bool*',
        },
    },
};
sub officess2csv {
    my %args = @_;

    my $input_file = $args{input_file} or return [400, "Please specify input_file"];
    my $output_file_or_dir = $args{output_file_or_dir} or return [400, "Please specify output_file_or_dir"];

    if (-e $output_file_or_dir && !$args{overwrite}) {
        return [412, "Output file/dir '$output_file_or_dir' already exists, not overwriting unless you specify --overwrite"];
    }

  CONVERT_TO_XLSX: {
        last if $input_file =~ /\.xlsx\z/i;
        require File::Copy;
        require File::Temp;
        require File::Which;
        require IPC::System::Options;

        my $libreoffice_path = $args{libreoffice_path} //
            File::Which::which("libreoffice") //
              File::Which::which("soffice");
        unless (defined $libreoffice_path) {
            log_debug "libreoffice is not in PATH, skipped trying to use libreoffice";
            last;
        }

        $input_file =~ /(.+)\.(\w+)\z/ or return [412, "Please supply input file with extension in its name (e.g. foo.doc instead of foo)"];
        my ($name, $ext) = ($1, $2);

        my $tempdir = File::Temp::tempdir(CLEANUP => !$ENV{DEBUG});
        my ($temp_fh, $temp_file) = File::Temp::tempfile(undef, SUFFIX => ".$ext", DIR => $tempdir);
        (my $temp_out_file = $temp_file) =~ s/\.\w+\z/.xlsx/;
        File::Copy::copy($input_file, $temp_file) or do {
            return [500, "Can't copy '$input_file' to '$temp_file': $!"];
        };
        log_debug "Converting $input_file -> $temp_out_file ...";
        IPC::System::Options::system(
            {die=>1, log=>1},
            $libreoffice_path, "--headless", "--convert-to", "xlsx", $temp_file, "--outdir", $tempdir);

        $input_file = $temp_out_file;
        log_trace "input xlsx file=$input_file";
    }

    #require Text::Iconv;
    #my $converter = Text::Iconv->new("utf-8", "windows-1251");
    my $converter;

    require Spreadsheet::XLSX;
    my $xlsx = Spreadsheet::XLSX->new($input_file, $converter);
    my @all_worksheets = map { $_->{Name} } @{ $xlsx->{Worksheet} };
    log_debug "Available worksheets in this workbook: %s", \@all_worksheets;
    my @worksheets;
    if ($args{all_worksheets}) {
        @worksheets = @all_worksheets;
        log_debug "Will be exporting all worksheet(s): %s", \@worksheets;
    } elsif ($args{worksheets}) {
        @worksheets = @{ $args{worksheets} };
        log_debug "Will be exporting these worksheet(s): %s", \@worksheets;
    } else {
        log_debug "Will only be exporting the first worksheet ($all_worksheets[0])";
        @worksheets = ($all_worksheets[0]);
    }

    my @output_files;
    if (@worksheets == 1 && !$args{always_dir}) {
        @output_files = ($output_file_or_dir);
    } else {
        unless (-d $output_file_or_dir) {
            log_debug "Creating directory $output_file_or_dir ...";
            mkdir $output_file_or_dir or do {
                return [500, "Can't mkdir $output_file_or_dir: $!, bailing out"];;
            };
        }
        for (@worksheets) {
            # XXX convert to safe filename
            push @output_files, "$output_file_or_dir/$_.csv";
        }
    }

    require Text::CSV_XS;
    my $csv = Text::CSV_XS->new({binary=>1});
  WRITE_WORKSHEET: {
        for my $i (0..$#worksheets) {
            my $worksheet = $worksheets[$i];
            my $output_file = $output_files[$i];
            log_debug "Outputting worksheet $worksheet to $output_file ...";

            my $sheet;
            for my $sheet0 (@{ $xlsx->{Worksheet} }) {
                if ($sheet0->{Name} eq $worksheet) {
                    $sheet = $sheet0; last;
                }
            }
            unless ($sheet) {
                log_error "Cannot find worksheet $worksheet, skipped";
                next WRITE_WORKSHEET;
            }

            open my $fh, ">", $output_file or do {
                log_error "Cannot open output file '$output_file': $!, skipped";
                next WRITE_WORKSHEET;
            };

            $sheet -> {MaxRow} ||= $sheet -> {MinRow};
            for my $row ($sheet->{MinRow} .. $sheet->{MaxRow}) {
                $sheet->{MaxCol} ||= $sheet->{MinCol};
                my @row;
                foreach my $col ($sheet->{MinCol} ..  $sheet->{MaxCol}) {
                    my $cell = $sheet->{Cells}[$row][$col];
                    push @row, $cell ? $cell->{Val} : undef;
                }
                $csv->combine(@row);
                print $fh $csv->string, "\n";
            }
        }
    }

    [200, "OK"];
}

1;
#ABSTRACT: Utilities related to Office suite files (.doc, .docx, .odt, .xls, .xlsx, .ods, etc)

=head1 DESCRIPTION

This distributions provides the following command-line utilities:

# INSERT_EXECS_LIST


=head1 ENVIRONMENT

=head2 DEBUG

If set to true, will not clean up temporary directories.


=head1 SEE ALSO

L<App::MSOfficeUtils>, L<App::LibreOfficeUtils>

=cut

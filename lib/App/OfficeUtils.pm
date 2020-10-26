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

1;
#ABSTRACT: Utilities related to Office suite files (.doc, .docx, .odt, .xls, .xlsx, .ods, etc)

=head1 DESCRIPTION

This distributions provides the following command-line utilities:

# INSERT_EXECS_LIST


=head1 SEE ALSO

L<App::MSOfficeUtils>, L<App::LibreOfficeUtils>

=cut

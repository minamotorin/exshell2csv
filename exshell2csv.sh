#!/bin/sh

####################### exshell2csv version 0.2 ########################
# exshell2csv: Small script to convert Excel to CSV, written in shell script only. No additional packages are required.
# Dependencies: Bourne Shell, sed, awk, and unzip.
# Usage: exshell2csv -h
# Reporitory: https://github.com/minamotorin/exshell2csv
# License: GNU General Public License Version 3 (https://www.gnu.org/licenses/gpl-3.0.html).
########################################################################

if [ "$1" = '-h' ]
then
  cat<<EOF
exshell2csv Version 0.2

Usage: `basename $0` [-h] [XLSX] [SHEET ID]

  -h               : Show this help
  [XLSX]           : Show list of SHEET IDS of XLSX file
  [XLSX] [SHEET ID]: Convert XLSX file’s SHEET ID to CSV (output to STDOUT)
EOF
exit
fi

if ! [ -e "$1" ]
then
  echo 'Error: file '"$1"' doesn’t exist' >&2
  exit 1
fi

if [ "$2" = '' ]
then
  unzip -p "$1" xl/workbook.xml                                        |
  sed '
    $!N
    H
    $!D
    x
    s/\n//g
    s/.*<sheets>//
    s/<\/sheets>.*//
    s/<sheet/\
&/g
  '                                                                    |
  sed 's/.* name="\(.*\)" sheetId="\([^"]*\)".*/\2: \1/; /^$/d'
  exit $?
fi

(
  unzip -p "$1" xl/sharedStrings.xml                                   |
  awk '{gsub("\\r", ""); print}'                                       |
  sed '
    1{
      s/^<?xml [^>]*>//
      /^$/d
      :loop
      /^[^<]/ s/.//
      t loop
    }
  '                                                                    |
  sed '
    1{/^<?xml/d;}
    s/<si>/\
&/g
    s/<si [^>]*>/\
<si>/g
  '                                                                    |
  sed '
    1d
    :loop
    /<\/si>/!{
      N
      bloop
    }

    :si
    /^<si>/! s/.//
    t si
    s/^<si>//

    s/\\/&&/g
    s/\n/\\n/g
    s/<\/si>.*//
  '                                                                    |
  sed '
    :topen
    /^<t>/!{ /^<t /!{
      s/^<[^>]*>//
      t topen
    }; }
    s/^<t>//
    s/^<t [^>]*>//

    H
    x
    s/<\/t>.*//
    s/\n//
    x

    :tclose
    /^<\/t>/!{
      s/.//
      t tclose
    }
    s/^<\/t>//

    /<\/t>/ b topen
    s/.*//
    x
    '                                                                  |
    sed 's/^/l /'

  echo

  unzip -p "$1" xl/worksheets/sheet"$2".xml                            |
  sed '
    s/<c [^>]*>/\
&\
/g
    s/<\/c>/\
&\
/g
  '                                                                    |
  sed -n '/<c .*[^/]>/, /<\/c>/p'                                      |
  sed '
    /^<c /{
      /t="s"/ s/.* r="\([^"]*\)" .*/\1 s/
      /^<c/ s/.* r="\([^"]*\)" .*s="\([0-9]*\)".*/\1 \2/
      /^<c/ s/.* r="\([^"]*\)">/\1 v/
      :loop
      N
      /<\/c>/! bloop
      s/\n/ /g
    }
    s/<\/\{0,1\}[vc]>//g
    s/<f\( [^>]*\)\{0,1\}>.*<\/f>//
    s/<f [^>]*\/>//
    s/^[A-Z]*/& /
  '
)                                                                      |
cat

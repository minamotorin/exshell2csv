#!/bin/sh

if [ "$1" = '' ] || [ "$1" = '-h' ]
then
  cat<<EOF >&2
Build script for exshell2csv (for version 0.1)
Usage: `basename $0` [exshell2csv] [CUSTOM] [FUNCTIONS]
  exshell2csv: output path (OVERWRITTEN)
  CUSTOM     : customize format (option)
  FUNCTION   : functions for customize (option)
EOF
  exit 1
fi

if ! [ -e exshell2csv.sh ] || ! [ -e exshell2csv.awk ]
then
  echo 'Error: required file doesn’t exist (exshell2csv.sh or exshell2csv.awk)' >&2
  exit 2
fi

cat exshell2csv.sh | sed 's/^cat/&                                                                    |/' >"$1"
echo "awk '" >>"$1"
if [ -e "$2" ]
then
  cat exshell2csv.awk | sed '/# CUSTOMIZE AREA #/r'"$2" | sed 's/^/  /' >>"$1"
else
  cat exshell2csv.awk | sed 's/^/  /' >>"$1"
fi
if [ -e "$3" ]
then
  echo >>"$1"
  cat "$3" >>"$1"
fi
echo "'" >>"$1"

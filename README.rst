README of exshell2csv, Version 0.1.1.

Abstract
########

`exshell2csv <https://github.com/minamotorin/exshell2csv>`_: Small script to convert Excel to CSV, written in shell script only. No additional packages are required.

Background
##########

I wanted convert Excel (``*.xlsx`` file) to CSV in command line. I found softwares and packages to do this but these sowtwares are too large.
Who want to install new packages just to convert? Why there is no small script to do that?
So I wrote this.

:Suggestion: **STOP Using EXCEL**

  :I want to write documents: Use `Markdown <https://docs.github.com/en/github/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax>`_.
  :I want to do something like a spreadsheet: Use `R <https://www.r-project.org/>`_.
  :I want to write documents with a spreadsheet: Use `R Markdown <https://rmarkdown.rstudio.com/>`_.
  :I surely have to use a spreadsheet: Use `(GNU) Emacs <https://www.gnu.org/software/emacs/>`_ `Org-mode <https://orgmode.org/>`_.

Installation
############

Clone `repository <https://github.com/minamotorin/exshell2csv>`_ and put ``exshell2csv`` on a ``PATH``.

And add permission: Run ``chmod +x /path/to/exshell2csv``.

Usage
#####

:Show Usage: Run ``exshell2csv`` or ``exshell2csv -h``.

:Show list of sheet ids and sheet names: Run ``exshell2csv /path/to/excel.xlsx``. Output is ``[SHEET ID]: [SHEET NAME]``.

:Convert Excel file’s sheet number [SHEET ID] to CSV: Run ``exshell2csv /path/to/excel.xlsx [SHEET ID]``. Not ``[SHEET NAMS]`` but ``[SHEET ID]``. Output to STDOUT.

If you are Microsoft Windows user, maybe you have to run ``exshell2csv [APGUMENTS] | awk "{gsub("$", "\\r"); print}"`` due to carriage ruturn difference. I’ve never checked if this code is required or not.

NOTE
####

I’ve never read documentation of Open Document Format. Some features will be wrong.

Some values will be not formated because some format features are not implemented. It is hard to check default format styles of Ecxcel, so there are only 2 format styles are supported.

And custom styles (defined in ``xl/styles.xml``) is also unsupported. I don’t know how to read ``xl/styles.xml``. (``xl/`` will be created by ``unzip /path/to/excel.xlsx``.)

Customize
*********

You can add format style yourself\:

1. Record cell number of line and column as ``[LINE]`` and ``[COLUMN]`` where value you want to format is on.

2. Run ``exshell2csv.sh /path/to/excel.xlsx [SHEET ID] | sed -n '/^l/d; /^[COLUMN] [LINE]/p'``. You have to replace arguments of exshell2csv.sh and ``[COLUMN]`` and ``[LINE]`` to 1.’s cell’s ones.

Output is ``[COLUMN] [LINE] [FORMAT ID] [VALUES]``. It means format style ``[FORMAT ID]`` will format ``[VALUE]`` as which you want.

3. Write format awk script in ``[CUSTOM FORMAT STYLE]`` file like this::

    else if ($3==[FORMAT ID]) {
      VALUES = after(4);
      VALUES = YourFormatScript(VALUES);
      cell[$2, $1] = VALUES;
    }

You can define new functions in ``[USER FUNCTIONS]`` file and use it. See Build_ section. Function ``after``, ``ALPH_advance``, ``ALPH_lt``, and ``fdate`` were already defined.

Dependencies
############

- ``/bin/sh`` (Bourne Shell)

  I don’t know does this script work on Ubuntu or not because Ubuntu’s ``/bin/sh`` is ``dash``.

- ``sed``

- ``awk`` (nawk or gawk)

- ``unzip``

Build
#####

This section shows how to make one script file from ``exshell2csv.sh``, ``exshell2csv.awk``, ``[CUSTOM FORMAT STYLE]`` (option), and ``[USER FUNCTIONS]`` (option).

Following commands are required run in same directory with ``exshell2csv.sh`` and ``exshell2csv.awk``.

``make.sh``’s first argument is output path. **WARNING**: If output path is exist, path will be *overwritten*.

:plain:

  Run ``make.sh [OUTPUT PATH]`` and make customizeless exshell2csv. ``[OUTPUT PATH]`` will be *overwritten*.

:test:

  You shold test if scripts work fine before build.
  
  Run ``sh exshell2csv.sh /path/to/excel.xlsx [SHEET ID] | awk -f exshell2csv.awk`` to test plain ``exshell2csv``. 

  For using custom format style, run followings::

    sed -e '/# CUSTOMIZE AREA #/r[CUSTOM FORMAT STYLE]' -e '$r[USER FUNCTIONS]' exshell2csv.awk > yourexshell2csv.awk
    sh exshell2csv.sh /path/to/excel.xlsx [SHEET ID] | awk -f yourexshell2csv.awk

  ``yourexshell2csv.awk`` will be overwritten.

:customize:

  If you want to add format style, run ``make.sh [OUTPUT PATH] [CUSTOM FORMAT STYLE]``.

  Or if you want to add functions, run ``make.sh [OUTPPUT PATH] [CUSTOM FORMAT STYLE] [USER FUNCTIONS]``.
  
  ``[OUTPUT PATH]`` will be *overwritten*.
  
  See also: Customize_ section in NOTE_ section.

Q&A
###

:Why couldn’t I use a sheet name to select the sheet?:

  Due to risk of a number sheet name.

:There are cells which have diference between original Excel and output CSV:

  CSV’s value on the cells are inner expression of Excel. The feature to format inner expression to string as same as Excel is not implemented. See NOTE_ section.

Reference
#########

Similar Projects
****************

There are many *softwares* or *packages* to convert Excel to CSV.

:`Microsoft Excel <https\://www.microsoft.com/en-us/microsoft-365/excel>`_:

  Excel can convert Excel file to CSV.

TODO
  Add similar projects and hyper links.

Issue
#####

If you have questions or feedbacks, or found bugs, typographical errors, wrong English or codes, or something else, pleas use `GitHub issue <https://github.com/minamotorin/exshell2csv/issues>`_ feel free.

Knowledge Bugs
**************

:leap year: Excel judges year 1900 is a leap year. But this script is not. This is Exces’s bug (due to compatibility). I didn’t implement this because I don’t know the details.

License
#######

This project is under the `GNU General Public License Version 3 <https://www.gnu.org/licenses/gpl-3.0.html>`_.

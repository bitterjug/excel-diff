Excel Textconv
===============

A script to render an excel (.xlsx) workbook as text suitable to use with
`diff` and, in particular, `git diff`, to get readable diffs, for example with
Git.  Inspired by William Usher's [`git_diff_xlsx`](https://github.com/willu47/git_diff_xlsx/) 
which he wrote for working with the UK TIMES model (see [his blog](https://wiki.ucl.ac.uk/display/~ucftpw2/2013/10/18/Using+git+for+version+control+of+spreadsheet+models+-+part+1+of+3)).
I wrote this in Ruby (using `rubyXL`) because thats the language of the other [UK-TIMES tools](https://github.com/decc/times-excel-tools)


Dependencies
-------------

- [RubyXL](https://github.com/weshatheleopard/rubyXL)


Features
--------

Options to:
- hide formulas, and show only values (useful for regression testing)
- hide calculated values for formulas and show just the formulas (useful for git diffs)


Install
--------

On Linux, I create a link called `xl-textconv` in my `~/bin`. 


Use
----

    Usage: xl-textconv [options] <workbook-file>
        -v, --no-values                  Exclude calculated values
        -f, --no-formulas                Exclude cell formulas
        -h, --help                       Show this message


Set up Git to use it
---------------------

See [the Git Book](http://git-scm.com/book/en/v2/Customizing-Git-Git-Attributes#Binary-Files)

Tell git that files ending `xlsx` are to be treated as Excel for diff:
create `.gitattributes` file in the root of your repo with the following:

    *.xlsx diff=excel

Tell Git how to diff Excel files: in `.git/config`:

    [diff "excel"]
        binary = True
        textconv = xl-textconv -v
        cachetextconv = true

See also [This node.js version](https://github.com/pismute/node-textconv) with
some interesting notes on use, including using
[cachetextconv](https://git-scm.com/docs/gitattributes/1.7.9)


Caveat emptor
--------------

- I'm not a Ruby programmer.
- There are no tests.

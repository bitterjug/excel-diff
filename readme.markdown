Excel Textconv
===============

A script to render an excel (.xlsx) workbook as text suitable to use with
`diff` and, in particular, `git diff`, to get readable diffs, for example with
Git.  Inspired by
[`git_diff_xlsx`](https://wiki.ucl.ac.uk/display/~ucftpw2/2013/10/18/Using+git+for+version+control+of+spreadsheet+models+-+part+1+of+3).

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


Set up Git to use it:
---------------------
See [the Git Book](http://git-scm.com/book/en/v2/Customizing-Git-Git-Attributes#Binary-Files)

Tell git that files ending `xlsx` are to be treated as Excel for diff:
create `.gitattributes` file in the root of your repo with the following:

    *.xlsx diff=excel

Tell Git how to diff Excel files: in `.git/config`:

    [diff "excel"]
        binary = True
        textconv = xl-textconv -v


    


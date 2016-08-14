ExtDiff [![Status Aquana][status-aquana]][andivionian-status-classifier]
=======

This is a small command line script that will compare two files using Microsoft
Word file comparison tool. Microsoft Word will be started using COM automation.

To run the script, execute it through PowerShell like this:

```console
$ powershell Diff-Word.ps1 oldfile.docx newfile.docx
```

Idea taken from [TortoiseSVN diff-doc script][tortoisesvn-diff-doc].

[andivionian-status-classifier]: https://github.com/ForNeVeR/andivionian-status-classifier#status-aquana-
[tortoisesvn-diff-doc]: https://sourceforge.net/p/tortoisesvn/code/27268/tree/trunk/contrib/diff-scripts/diff-doc.js

[status-aquana]: https://img.shields.io/badge/status-aquana-yellowgreen.svg

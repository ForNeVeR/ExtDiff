ExtDiff [![Status Aquana][status-aquana]][andivionian-status-classifier]
=======

This is a small command line script that will compare two files using Microsoft
Word file comparison tool. Microsoft Word will be started using COM automation.

It is useful as a diff tool for Word-related file types.

## Using via command line

To run the script, execute it through PowerShell like this:

```console
$ powershell -File Diff-Word.ps1 oldfile.docx newfile.docx
```

Or via the batch file:

```console
$ diff-word.cmd oldfile.docx newfile.docx
```

## Using via Git Integration

You can also use this tool with git, so that `git diff` will use Microsoft Word
to diff `*.docx` files.

To do this, you must configure your `.gitattributes` and `.gitconfig` to support
a custom diff tool.

To configure your `.gitattributes`, open or create a file called
`.gitattributes` in your git repo's root directory. Add the following text to a
new line in this file:

```
*.docx diff=word
```

To configure your `.gitconfig`, open or create the file in your home directory.
Then, add the following to your `.gitconfig`:

```ini
[diff "word"]
	command = <pathToExtDiffFolder>/diff-word-wrapper.cmd
```

Replace `<pathToExtDiffFolder>` with the path to this repo's
location on disk.

-------

Idea taken from [TortoiseSVN diff-doc script][tortoisesvn-diff-doc].

[andivionian-status-classifier]: https://github.com/ForNeVeR/andivionian-status-classifier#status-aquana-
[tortoisesvn-diff-doc]: https://sourceforge.net/p/tortoisesvn/code/27268/tree/trunk/contrib/diff-scripts/diff-doc.js

[status-aquana]: https://img.shields.io/badge/status-aquana-yellowgreen.svg

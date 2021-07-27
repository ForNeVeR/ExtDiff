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

## Using GUI form

To run the GUI form run file `start-gui.cmd`: it will open form that stays on top off all windows and you can drag files from Windows Explorer to two text fields (first for old document and second for revised document). When you click "Compare" button it will start Word application with chosen documents as in command line usage and this form will loose "stay on top of all windows" property. It will regain this property after button "Clear" is clicked.

## Using via Git Integration

You can also use this tool with git, so that `git diff` will use Microsoft Word
to diff `*.docx` files.

To do this, you must configure your `.gitattributes` and `.gitconfig` to support
a custom diff tool.

### `.gitattributes`

To configure your `.gitattributes`, open or create a file called
`.gitattributes` in your git repo's root directory. Add the following text to a
new line in this file:

```
*.docx diff=word
```

It is also possible to create a global `.gitattributes` file that will be
applied to every repository in a system. To do that, create a file
`.gitattributes` in your home directory, and then perform the following command:

```console
git config --global core.attributesfile ~/.gitattributes
```

### `.gitconfig`

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

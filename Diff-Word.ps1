param(
    $BaseFileName,
    $ChangedFileName
)

# Constants
$wdCompareTargetNew = 2

$word = New-Object -ComObject Word.Application
$word.Visible = $true
$document = $word.Documents.Open($BaseFileName, $true, $true)
$document.Compare($ChangedFileName, [ref]"Comparison", [ref]$wdCompareTargetNew, [ref]$true, [ref]$true)

$word.ActiveDocument.Saved = 1
#$wdDoNotSaveChanges = 0
#$document.Close(wdDoNotSaveChanges)

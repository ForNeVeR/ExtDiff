param(
    $BaseFileName,
    $ChangedFileName
)

# Constants
$wdDoNotSaveChanges = 0
$wdCompareTargetNew = 2

try {
	$word = New-Object -ComObject Word.Application
	$word.Visible = $true
	$document = $word.Documents.Open($BaseFileName, $false, $true)
	$document.Compare($ChangedFileName, [ref]"Comparison", [ref]$wdCompareTargetNew, [ref]$true, [ref]$true)

	$word.ActiveDocument.Saved = 1

	# Now close the document so only compare results window persists:
	$document.Close([ref]$wdDoNotSaveChanges)
} catch {
	Add-Type -AssemblyName System.Windows.Forms
	[System.Windows.Forms.MessageBox]::Show($_.Exception)
}

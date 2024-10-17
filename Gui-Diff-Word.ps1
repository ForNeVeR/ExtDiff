$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object System.Windows.Forms.Form
$Form.ClientSize = New-Object System.Drawing.Point(595, 162)
$Form.text = 'Compare MS Word Documents (drop files here)'
#Make a form topmost window - good for drag and drop operations
$Form.TopMost = $true
$Form.FormBorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$olddoc = New-Object System.Windows.Forms.TextBox
$olddoc.multiline = $true
$olddoc.width = 580
$olddoc.height = 50
$olddoc.location = New-Object System.Drawing.Point(6, 8)
$olddoc.AllowDrop = $true

$newdoc = New-Object System.Windows.Forms.TextBox
$newdoc.multiline = $true
$newdoc.width = 580
$newdoc.height = 50
$newdoc.location = New-Object System.Drawing.Point(6, 65)
$newdoc.AllowDrop = $true

$cmpbtn = New-Object System.Windows.Forms.Button
$cmpbtn.text = "Compare"
$cmpbtn.width = 446
$cmpbtn.height = 30
$cmpbtn.location = New-Object System.Drawing.Point(140, 126)

$clrbtn = New-Object System.Windows.Forms.Button
$clrbtn.text = "Clear"
$clrbtn.width = 120
$clrbtn.height = 30
$clrbtn.location = New-Object System.Drawing.Point(6, 126)

$Form.controls.AddRange(@($olddoc, $newdoc, $cmpbtn, $clrbtn))

$olddoc.Add_DragDrop( { DnD $this $_ })
$olddoc.Add_DragOver( { DnO $this $_ })
$newdoc.Add_DragDrop( { DnD $this $_ })
$newdoc.Add_DragOver( { DnO $this $_ })
$cmpbtn.Add_Click( { DoCompare $this $_ })
$clrbtn.Add_Click( { ClearInputs $this $_ })

function DnO ($evSource, $evtArgs) {
    if ($evtArgs.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $evtArgs.Effect = [Windows.Forms.DragDropEffects]::Copy
    }
    else {
        $evtArgs.Effect = [Windows.Forms.DragDropEffects]::None
    }
    $Form.ActiveControl = $cmpbtn
}
function DnD ($evSource, $evtArgs) {
    
    $files = $evtArgs.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    if ($files.Count -eq 1)
    {
        $evSource.text = $files[0]
    } 
    elseif ($files.Count -gt 1) 
    {
        if ($evSource -eq $olddoc)
        {
            $olddoc.text = $files[0]
            $newdoc.text = $files[1]
        }
        else 
        {
            $olddoc.text = $files[1]
            $newdoc.text = $files[0]
        }
    }
}
function DoCompare ($evSource, $evArgs) {
    $Form.TopMost = $false
    $BaseFileName = $olddoc.text
    $ChangedFileName = $newdoc.text
    # Remove the readonly attribute because Word is unable to compare readonly
    # files:
    $baseFile = Get-ChildItem $BaseFileName
    if ($baseFile.IsReadOnly) {
        $baseFile.IsReadOnly = $false
    }
    # Constants
    $wdDoNotSaveChanges = 0
    $wdCompareTargetNew = 2
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $document = $word.Documents.Open($BaseFileName, $false, $false)
        $document.Compare($ChangedFileName, [ref]"Comparison", [ref]$wdCompareTargetNew, [ref]$true, [ref]$true)

        $word.ActiveDocument.Saved = 1

        # Now close the document so only compare results window persists:
        $document.Close([ref]$wdDoNotSaveChanges)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception)
    }
}
function ClearInputs ($evSource, $evtArgs) {
    $olddoc.text = ""
    $newdoc.text = ""
    $Form.TopMost = $true
}

[void]$Form.ShowDialog()

$Filename='D:\python\test\L16_connectivity_template.doc'
$Word = New-Object -ComObject Word.Application
$Word.Visible = $false
$Word.DisplayAlerts = 'wdAlertsNone'
$Doc = $Word.Documents.Open($Filename)

$Row = $Doc.Tables.Item(1).Cell(2,1)
Write-Host $Row.Range.Text

$Row = $Doc.Tables.Item(1).Cell(3,1)
Write-Host $Row.Range.Text


$Word.Quit()
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable Word 
param (
    [string]$file
)
Add-Type -Path .\itextsharp.dll
$cwd = Get-Location

$pdf = $cwd.Path + "\DocDjango_SP92_1.0.0_28pkgs_27avril2017.pdf"

$reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $file

Write-Output "Item Name;Obsolete;Uninstall Prev;Install Dep;Runtime Dep;Install Nodes;Accs Stop;Server Reboot;Install Duration"
for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
    $strategy = new-object  'iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy'
    $currentText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page, $strategy);
    #[string[]]$Text += [system.text.Encoding]::UTF8.GetString([System.Text.ASCIIEncoding]::Convert( [system.text.encoding]::default  , [system.text.encoding]::UTF8, [system.text.Encoding]::Default.GetBytes($currentText)));
    [string[]]$currentLine = [system.text.Encoding]::UTF8.GetString([System.Text.ASCIIEncoding]::Convert( [system.text.encoding]::default  , [system.text.encoding]::UTF8, [system.text.Encoding]::Default.GetBytes($currentText)));
    foreach ($line in $currentLine.Split([char]0x000A)) {
        if ($line -like "*which are obsolete*") {
            $wr = $line
            $itemName = $prevLine
            $add = 1
        } ElseIf ($line -like "*Comment*") {
            $startIndex = $wr.IndexOf(":")
            $startIndex += 1
            $obsolete = $wr.Substring($startIndex)
            $add = 0
        } ElseIf ($line -like "*Uninstall previous Version*") {
            $startIndex = $line.IndexOf(":")
            $startIndex += 1
            $uninstallPrev= $line.Substring($startIndex)
        } ElseIf ($line -like "*Installation Dependencies*") {
            $wr = $line
            $add = 1
        } ElseIf ($line -like "*Runtime Dependencies*") {
            $startIndex = $wr.IndexOf(":")
            $startIndex += 1
            $installDep = $wr.Substring($startIndex)
            $wr = $line
        } ElseIf ($line -like "*PTR(s)*") {
            $startIndex = $wr.IndexOf(":")
            $startIndex += 1
            $runtimeDep = $wr.Substring($startIndex)
            $add = 0
        } ElseIf ($line -like "*Installation Nodes*") {
            $startIndex = $line.IndexOf(":")
            $startIndex += 1
            $installNodes= $line.Substring($startIndex)
        } ElseIf ($line -like "*Require ACCS System Stop*") {
            $startIndex = $line.IndexOf(":")
            $startIndex += 1
            $accsStop= $line.Substring($startIndex)
        } ElseIf ($line -like "*Require Server reboot*") {
            $startIndex = $line.IndexOf(":")
            $startIndex += 1
            $serverReboot= $line.Substring($startIndex)
        } ElseIf ($line -like "*Installation Duration*") {
            $startIndex = $line.IndexOf(":")
            $startIndex += 1
            $installDuration= $line.Substring($startIndex)
            $wr = $itemName + ";" + $obsolete + ";" + $uninstallPrev + ";" + $installDep + ";" + $runtimeDep + ";" + $installNodes + ";" + $accsStop + ";" + $serverReboot + ";" + $installDuration
            Write-Output $wr
            $itemName = ""
            $obsolete = ""
            $uninstallPrev = ""
            $installDep = ""
            $runtimeDep = ""
            $installNodes = ""
            $accsStop = ""
            $serverReboot = "" 
            $installDuration = ""
        } ElseIf($add) {
            $wr += $line
        } Else {
            $prevLine = $line    
        }        
        
    }
    #Write-Output $currentText
}
#Write-Output $Text

$reader.Close()
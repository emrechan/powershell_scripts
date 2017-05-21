param (
    [string]$file
)
Add-Type -Path .\itextsharp.dll
$cwd = Get-Location

$pdf = $cwd.Path + "\1058DS0837_SP90_1.0.1_Rev-.pdf"

$reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $file


for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
    $strategy = new-object  'iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy'
    $currentText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page, $strategy);
    #[string[]]$Text += [system.text.Encoding]::UTF8.GetString([System.Text.ASCIIEncoding]::Convert( [system.text.encoding]::default  , [system.text.encoding]::UTF8, [system.text.Encoding]::Default.GetBytes($currentText)));
    [string[]]$currentLine = [system.text.Encoding]::UTF8.GetString([System.Text.ASCIIEncoding]::Convert( [system.text.encoding]::default  , [system.text.encoding]::UTF8, [system.text.Encoding]::Default.GetBytes($currentText)));
    foreach ($line in $currentLine.Split([char]0x000A)) {
        if ($line -like "*which are obsolete*") {
            $startIndex = $line.IndexOf(":")
            $startIndex += 1
            $wr = $prevLine + ";" + $line.Substring($startIndex)
            $add = True
        } ElseIf ($line -like "*Installation Nodes*") {
            $startIndex = $line.IndexOf(":")
            $startIndex += 1
            $wr += ";" + $line.Substring($startIndex)
        } ElseIf ($line -like "*Require ACCS System Stop*") {
            $startIndex = $line.IndexOf(":")
            $startIndex += 1
            $wr += ";" + $line.Substring($startIndex)
        } ElseIf ($line -like "*Require Server reboot*") {
            $startIndex = $line.IndexOf(":")
            $startIndex += 1
            $wr += ";" + $line.Substring($startIndex)
        } ElseIf ($line -like "*Installation Duration*") {
            $startIndex = $line.IndexOf(":")
            $startIndex += 1
            $wr += ";" + $line.Substring($startIndex)
            Write-Output $wr
        } Else {
            $prevLine = $line    
        }        
        
    }
    #Write-Output $currentText
}
#Write-Output $Text

$reader.Close()
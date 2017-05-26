param (
    [string]$file
)

#$Filename="D:\SVN\AccsRepo\TC\PT86704_-_AppendixB_Redlines2017\Interfaces\L16\T_VOL7_ES_SYST_LINK16_UNIT_001-001.doc"
if (Test-Path $file) {
  $Word = New-Object -ComObject Word.Application
  $Word.Visible = $false
  $Word.DisplayAlerts = $false
  $Doc = $Word.Documents.Open($file)

  $location = ([regex]::matches($Doc.Paragraphs.Item(7).Range.Text, "located at \w+")| %{$_.value})
  $location = $location -replace "located at "

  $ht = @{}

  for ($i = 3; $i -le $Doc.Tables.Item(2).Rows.Count; $i++) {
      try {
          if ($Doc.Tables.Item(2).Cell($i,5).Range.Text -match 'ES[D]{0,1}\d+-?\w{0,2}-?\w{0,2}/?\d?-?') {
              ([regex]::matches($Doc.Tables.Item(2).Cell($i,5).Range.Text,'ES\w?\d+-?\w{0,2}-?\w{0,2}/?\d?-?')) | % {$ht.Add($_.value, 1)}
          }
      } catch {

      }
  }
  $ht.Keys | % {
    $output = $file + ";" + $location + ";" + $_
    Write-Output $output
  }

  $Word.Quit()
  $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Word)
  [gc]::Collect()
  [gc]::WaitForPendingFinalizers()
  Remove-Variable Word
  Remove-Variable ht
}

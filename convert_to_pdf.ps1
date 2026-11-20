$word = New-Object -ComObject Word.Application
$word.Visible = $false
$docx = (Resolve-Path "DeccanPulse_Internship_Report.docx").Path
$pdf = [System.IO.Path]::ChangeExtension($docx, ".pdf")
if (Test-Path $pdf) { Remove-Item $pdf -Force }
$doc = $word.Documents.Open($docx)
$doc.SaveAs([ref]$pdf, [ref]17)
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Host "Word conversion done: $pdf"

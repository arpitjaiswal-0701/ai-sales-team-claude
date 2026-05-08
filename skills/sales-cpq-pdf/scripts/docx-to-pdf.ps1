# docx-to-pdf.ps1 — Convert a .docx to PDF using Word COM automation (Windows only).
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File docx-to-pdf.ps1 -InputDocx "C:\path\file.docx"
#   powershell -ExecutionPolicy Bypass -File docx-to-pdf.ps1 -InputDocx "C:\path\file.docx" -OutputPdf "C:\path\out.pdf"

param(
    [Parameter(Mandatory=$true)]
    [string]$InputDocx,
    [string]$OutputPdf = ""
)

$InputDocx = (Resolve-Path $InputDocx).Path

if ($OutputPdf -eq "") {
    $OutputPdf = [System.IO.Path]::ChangeExtension($InputDocx, ".pdf")
} else {
    $OutputPdf = [System.IO.Path]::GetFullPath($OutputPdf)
}

Write-Host "Converting: $InputDocx"
Write-Host "       To:  $OutputPdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false

try {
    $doc = $word.Documents.Open($InputDocx, [ref]$false, [ref]$true)  # ReadOnly=$true
    $doc.SaveAs2($OutputPdf, 17)   # 17 = wdFormatPDF
    $doc.Close([ref]$false)
    Write-Host "Done. PDF: $OutputPdf"
    Write-Host "Size: $([math]::Round((Get-Item $OutputPdf).Length / 1KB)) KB"
} catch {
    Write-Error "Conversion failed: $_"
    exit 1
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

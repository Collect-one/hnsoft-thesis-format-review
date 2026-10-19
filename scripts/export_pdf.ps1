param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [Parameter(Mandatory = $true)][string]$Pdf
)

$ErrorActionPreference = "Stop"
$docxPath = (Resolve-Path -LiteralPath $Docx).Path
if (-not (Test-Path -LiteralPath $docxPath -PathType Leaf)) { throw "DOCX not found: $docxPath" }
$pdfPath = [System.IO.Path]::GetFullPath($Pdf)
$pdfDir = Split-Path -Parent $pdfPath
if (-not (Test-Path -LiteralPath $pdfDir)) { New-Item -ItemType Directory -Force -Path $pdfDir | Out-Null }

$dir = Split-Path -Parent $docxPath
$leaf = Split-Path -Leaf $docxPath
$locks = @(
    (Join-Path $dir ("~$" + $leaf)),
    (Join-Path $dir ("~$" + $leaf.Substring([Math]::Min(2, $leaf.Length))))
) | Select-Object -Unique
foreach ($lock in $locks) {
    if (Test-Path -LiteralPath $lock) {
        throw "Word lock file exists: $lock. Save and close the document in Word first."
    }
}

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($docxPath, $false, $true, $false)
    $doc.SaveAs2($pdfPath, 17)  # wdFormatPDF
    $doc.Close(0)
    $doc = $null
    $word.Quit()
    $word = $null
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close(0) } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        $doc = $null
    }
    if ($word -ne $null) {
        try { $word.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        $word = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (-not (Test-Path -LiteralPath $pdfPath)) { throw "PDF was not created: $pdfPath" }
Write-Output ("PDF: " + $pdfPath + " (" + (Get-Item -LiteralPath $pdfPath).Length + " bytes)")

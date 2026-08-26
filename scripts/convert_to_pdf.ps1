# Converts a PPTX file to PDF using PowerPoint's own rendering engine (COM automation).
#
# LibreOffice-based conversion can fail to render OMML (PowerPoint-native equation
# objects) and other elements correctly, causing formulas/symbols to disappear.
# This script uses the PowerPoint application installed on Windows instead, so the
# PDF matches what is shown while editing.
#
# Run this script on the Windows side (PowerShell). From WSL, call it via
# convert_to_pdf.sh, which also falls back to LibreOffice when PowerPoint is
# unavailable.
#
# NOTE: keep this file plain ASCII. Windows PowerShell 5.1 (powershell.exe) reads
# .ps1 files without a BOM using the system codepage rather than UTF-8, so non-ASCII
# characters (e.g. Japanese comments) can be misread and corrupt script parsing.

param (
    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [string]$OutputFile
)

# Use .ProviderPath (not .Path) to get a plain filesystem path. For UNC paths such as
# \\wsl.localhost\..., PathInfo.Path can come back prefixed with a PowerShell provider
# qualifier ("Microsoft.PowerShell.Core\FileSystem::..."), which the PowerPoint COM API
# cannot parse and fails with "path not found" (HRESULT 0x80070003).
$pptxPath = (Resolve-Path $InputFile).ProviderPath

if ($OutputFile) {
    $pdfPath = [System.IO.Path]::GetFullPath($OutputFile)
} else {
    $pdfPath = [System.IO.Path]::ChangeExtension($pptxPath, ".pdf")
}

$ppt = $null
$presentation = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application

    # On some PowerPoint versions/environments, Application.Visible cannot be set to
    # msoFalse (unlike Word/Excel) and the assignment itself throws. This failure does
    # not affect the conversion, so it is safe to ignore.
    try {
        $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    } catch {
        Write-Host "Note: could not hide the PowerPoint window (continuing): $_"
    }

    $presentation = $ppt.Presentations.Open(
        $pptxPath,
        [Microsoft.Office.Core.MsoTriState]::msoTrue,
        [Microsoft.Office.Core.MsoTriState]::msoFalse,
        [Microsoft.Office.Core.MsoTriState]::msoFalse
    )

    # ppSaveAsPDF = 32
    $presentation.SaveAs($pdfPath, 32)

    Write-Host "Converted successfully: $pdfPath"
}
catch {
    Write-Error "Failed to convert: $_"
    exit 1
}
finally {
    if ($presentation) {
        $presentation.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }

    if ($ppt) {
        $ppt.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

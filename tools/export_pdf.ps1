# dist/ 안의 docx 를 전부 같은 이름의 pdf 로 내보낸다. (Word 필요)
#   powershell -ExecutionPolicy Bypass -File tools\export_pdf.ps1
$dist = Join-Path $PSScriptRoot '..\dist' | Resolve-Path
$w = New-Object -ComObject Word.Application
$w.Visible = $false
try {
    Get-ChildItem -LiteralPath $dist -Filter *.docx |
        Where-Object { $_.Name -notlike '~$*' } |
        ForEach-Object {
            $out = [System.IO.Path]::ChangeExtension($_.FullName, '.pdf')
            $doc = $w.Documents.Open($_.FullName, $false, $true)
            $doc.SaveAs([ref]$out, [ref]17)
            $doc.Close([ref]0)
            Write-Output ("PDF  " + $_.BaseName)
        }
} finally {
    $w.Quit()
}

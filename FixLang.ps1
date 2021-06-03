Set-StrictMode -Version Latest

$version = "2021-06-03"
Write-Output "FixLang - built $version by Shinichi Akiyama"

function Write-Message ([string]$message) {
    Write-Output $message | Out-File $logfile -Append
    Write-Output $message
}

function backupFile ($path) {
    $directory = Split-Path $path
    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($path)
    $extension = $path.Extension
    $backup = Join-Path $directory "$fileNameWithoutExtension - backup$extension"
    $num = 2
    while (Test-Path $backup) {
        $backup = Join-Path $directory "$fileNameWithoutExtension - backup ($num)$extension"
        $num = $num + 1
    }
    Copy-Item $path $backup
    $backup
}

function ChangeLangID ($textRange) {
    foreach ($run in $textRange.Runs()) {
        $text = $run.Text -replace "[\x0b\x0d]", "`n"
        $languageID = $run.LanguageID
        $run.LanguageID = [Microsoft.Office.Core.MsoLanguageID]::msoLanguageIDJapanese
        if (($text -ne "") -and ($null -ne $languageID) -and ($run.LanguageID -ne $languageID)) {
            Write-Message "[$text] LanguageID has been changed from $languageID to Japanese"
        }
    }
}

function treatShape ($shape) {
    if ($shape.HasTextFrame -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
        if ($shape.TextFrame.HasText) {
            ChangeLangID $shape.TextFrame.TextRange
        }
    }
    elseif ($shape.HasTable -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
        foreach ($row in $shape.Table.rows) {
            foreach ($cell in $row.cells) {
                ChangeLangID $cell.shape.TextFrame.TextRange
            }
        }
    }
    elseif ($shape.GroupItems) {
        foreach ($item in $shape.GroupItems) {
            treatShape $item
        }
    }
}

Add-Type -AssemblyName Office
$app = New-Object -ComObject PowerPoint.Application
Write-Output "PowoerPoint version $($app.version)"

foreach ($file in Get-ChildItem -Path $args) {
    $logfile = [System.IO.Path]::GetDirectoryName($file) + "\" + [System.IO.Path]::GetFileNameWithoutExtension($file) + ".log"

    $backup = backupFile $file
    Write-Message "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") $file was backed up to $backup."

    $presentation = $app.Presentations.Open($file)
    $app.WindowState = [Microsoft.Office.Interop.PowerPoint.PpWindowState]::ppWindowMinimized
    Write-Message "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") $file was opened."

    foreach ($slide in $presentation.Slides) {
        Write-Message "--- Slide $($slide.SlideIndex) ---"
        foreach ($shape in $slide.Shapes) {
            treatShape $shape
        }
    }

    $presentation.Save()
    Write-Message "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") $file was saved."
    $presentation.Close()
}
Write-Output "All files were processed."

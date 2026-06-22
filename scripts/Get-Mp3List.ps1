# Get-Mp3List.ps1
# 指定フォルダ内の MP3 ファイル一覧を CSV で出力する（補助用）
#
# 使い方:
#   .\Get-Mp3List.ps1 -Mp3Folder "H:\MP3ライブラリ" -OutputCsv "mp3_list.csv"
#
# 引数:
#   -Mp3Folder : MP3 が入っているフォルダのパス（必須）
#   -OutputCsv : 出力 CSV ファイルパス（省略時: mp3_list.csv）

param(
    [Parameter(Mandatory = $true, HelpMessage = "MP3 フォルダのパスを指定してください")]
    [string]$Mp3Folder,

    [string]$OutputCsv = "mp3_list.csv"
)

if (-not (Test-Path $Mp3Folder)) {
    Write-Error "フォルダが見つかりません: $Mp3Folder"
    exit 1
}

Write-Host "MP3 ファイルを検索中: $Mp3Folder"
Write-Host "（大量ファイルがある場合は時間がかかります）"

$shell = New-Object -ComObject Shell.Application

$results = Get-ChildItem -Path $Mp3Folder -Filter "*.mp3" -Recurse -File |
ForEach-Object {
    $file = $_

    $folderObj = $shell.Namespace($file.DirectoryName)
    $fileObj   = $folderObj.ParseName($file.Name)
    $duration  = $folderObj.GetDetailsOf($fileObj, 27)

    if ([string]::IsNullOrWhiteSpace($duration)) {
        $duration = ""
    }

    [PSCustomObject]@{
        FileName     = $file.Name
        FullPath     = $file.FullName
        Folder       = $file.DirectoryName
        SizeMB       = [Math]::Round($file.Length / 1MB, 2)
        Duration     = $duration
        DateModified = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
    }
}

if (-not $results) {
    Write-Warning "MP3 ファイルが見つかりませんでした: $Mp3Folder"
    exit 0
}

$results | Export-Csv -Path $OutputCsv -Encoding UTF8 -NoTypeInformation

$count = ($results | Measure-Object).Count
Write-Host "完了: $count 件の MP3 ファイルを出力しました → $OutputCsv"

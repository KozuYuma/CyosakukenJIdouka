# Get-WavList.ps1
# NUENDO プロジェクトの Audio フォルダ内にある WAV ファイル一覧を CSV で出力する
#
# 使い方:
#   .\Get-WavList.ps1 -AudioFolder "H:\Project\Audio" -OutputCsv "wav_list.csv"
#
# 引数:
#   -AudioFolder : WAV が入っているフォルダのパス（必須）
#   -OutputCsv   : 出力 CSV ファイルパス（省略時: wav_list.csv）

param(
    [Parameter(Mandatory = $true, HelpMessage = "Audio フォルダのパスを指定してください")]
    [string]$AudioFolder,

    [string]$OutputCsv = "wav_list.csv"
)

# フォルダの存在確認
if (-not (Test-Path $AudioFolder)) {
    Write-Error "フォルダが見つかりません: $AudioFolder"
    exit 1
}

Write-Host "WAV ファイルを検索中: $AudioFolder"

# Shell.Application を使ってメタデータ（Duration）を取得する
$shell = New-Object -ComObject Shell.Application

$results = Get-ChildItem -Path $AudioFolder -Filter "*.wav" -Recurse -File |
ForEach-Object {
    $file = $_

    # Duration を Shell API から取得（列番号 27 = Duration）
    $folderObj = $shell.Namespace($file.DirectoryName)
    $fileObj   = $folderObj.ParseName($file.Name)
    $duration  = $folderObj.GetDetailsOf($fileObj, 27)

    # Duration が空の場合は空文字
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
    Write-Warning "WAV ファイルが見つかりませんでした: $AudioFolder"
    exit 0
}

# CSV 出力（UTF-8 BOM なし）
$results | Export-Csv -Path $OutputCsv -Encoding UTF8 -NoTypeInformation

$count = ($results | Measure-Object).Count
Write-Host "完了: $count 件の WAV ファイルを出力しました → $OutputCsv"

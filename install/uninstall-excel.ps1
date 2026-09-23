<#
韵词字幕 · Excel 字幕工具箱 - 卸载脚本

用法：右键这个文件 → "使用 PowerShell 运行"

会做的事：
  1. 删掉 install-excel.ps1 当初写进注册表的那条"开机自动加载"记录
  2. 删掉复制到 %APPDATA%\Microsoft\AddIns\ 下的 ExcelToSRT.xlam 和 reader.html

跑之前建议先关掉 Excel。
#>

$ErrorActionPreference = "Stop"

if (Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue) {
    Write-Host "检测到 Excel 正在运行。" -ForegroundColor Yellow
    Write-Host "删除文件这一步可能会因为文件被占用而失败，建议先关掉 Excel 再继续。"
    $answer = Read-Host "要现在继续吗？(y/n，直接回车默认 n)"
    if ($answer -ne "y") {
        Write-Host "已取消，请关掉 Excel 后重新运行这个脚本。"
        exit
    }
}

$installDir = Join-Path $env:APPDATA "Microsoft\AddIns"
$xlamTarget = Join-Path $installDir "ExcelToSRT.xlam"
$htmlTarget = Join-Path $installDir "reader.html"

Write-Host "===== 清理注册表里的自动加载记录 =====" -ForegroundColor Cyan

$officeVersions = @("16.0", "15.0", "14.0")
$removed = $false
foreach ($v in $officeVersions) {
    $excelOptionsKey = "HKCU:\Software\Microsoft\Office\$v\Excel\Options"
    if (-not (Test-Path $excelOptionsKey)) { continue }

    $item = Get-Item $excelOptionsKey
    foreach ($name in $item.GetValueNames()) {
        if ($name -match "^OPEN\d*$") {
            $value = (Get-ItemProperty -Path $excelOptionsKey -Name $name).$name
            if ($value -and ($value -like "*ExcelToSRT.xlam*")) {
                Remove-ItemProperty -Path $excelOptionsKey -Name $name
                Write-Host ("  已删除：{0}\{1}" -f $excelOptionsKey, $name)
                $removed = $true
            }
        }
    }
}
if (-not $removed) {
    Write-Host "  没找到对应的自动加载记录（可能已经手动删过了）"
}

Write-Host "===== 删除加载项文件 =====" -ForegroundColor Cyan

foreach ($f in @($xlamTarget, $htmlTarget)) {
    if (Test-Path $f) {
        Remove-Item $f -Force
        Write-Host ("  已删除：{0}" -f $f)
    }
}

Write-Host ""
Write-Host "卸载完成，重新打开 Excel 后「字幕工作」选项卡应该就不见了。" -ForegroundColor Green

<#
韵词字幕 · Excel 字幕工具箱 - 用户安装脚本

用法：
  把这个脚本和 ExcelToSRT.xlam、reader.html 放在同一个目录下，
  右键这个文件 → "使用 PowerShell 运行"（或者在 PowerShell 里 cd 到这个目录，
  运行 .\install-excel.ps1）

装的是什么：
  1. 把 ExcelToSRT.xlam 和 reader.html 复制到 Excel 默认的加载项目录
     （%APPDATA%\Microsoft\AddIns）
  2. 在当前 Windows 账户的注册表里加一条"开机自动加载此加载项"的记录
     —— 这和你手动在"文件→选项→加载项"里勾选是完全一样的效果，
     只是脚本帮你做了，不用再点那几步

装完之后，重新打开 Excel，功能区应该会自动出现"字幕工作"选项卡，
不需要再手动去加载项设置里勾选。

只影响当前 Windows 账户（不需要管理员权限，不会影响电脑上其他账户）。
如果想卸载：把 install-excel.ps1 目录下多出的 uninstall-excel.ps1 跑一下就行
（删除文件 + 删掉那条注册表记录）。

免责声明：这个脚本没有在真实 Windows + Excel 环境跑过验证（开发环境是
Linux 容器）。核心机制（复制文件到 AddIns 目录 + 写注册表 OPEN 键）是
Excel 加载项的标准做法，原理上没问题，但第一次跑如果报错，把报错信息
发回去，照着调。
#>

$ErrorActionPreference = "Stop"

if (Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue) {
    Write-Host "检测到 Excel 正在运行。" -ForegroundColor Yellow
    Write-Host "复制文件这一步可能会因为文件被占用而失败，建议先关掉 Excel（包括后台进程）再继续。"
    $answer = Read-Host "要现在继续吗？(y/n，直接回车默认 n)"
    if ($answer -ne "y") {
        Write-Host "已取消，请关掉 Excel 后重新运行这个脚本。"
        exit
    }
}

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlamSource = Join-Path $scriptDir "ExcelToSRT.xlam"
$htmlSource = Join-Path $scriptDir "reader.html"

if (-not (Test-Path $xlamSource)) {
    throw "找不到 ExcelToSRT.xlam，请确认它和 install-excel.ps1 放在同一目录下：$xlamSource"
}
if (-not (Test-Path $htmlSource)) {
    throw "找不到 reader.html，请确认它和 install-excel.ps1 放在同一目录下：$htmlSource"
}

Write-Host "===== 第 1 步：复制文件到 Excel 加载项目录 =====" -ForegroundColor Cyan

$installDir = Join-Path $env:APPDATA "Microsoft\AddIns"
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
}

$xlamTarget = Join-Path $installDir "ExcelToSRT.xlam"
$htmlTarget = Join-Path $installDir "reader.html"

Copy-Item $xlamSource $xlamTarget -Force
Copy-Item $htmlSource $htmlTarget -Force

Write-Host ("  已复制到：{0}" -f $installDir)

Write-Host "===== 第 2 步：注册开机自动加载 =====" -ForegroundColor Cyan

# 找到本机已安装的 Excel 对应的注册表版本号
$officeVersions = @("16.0", "15.0", "14.0")
$excelOptionsKey = $null
foreach ($v in $officeVersions) {
    $probe = "HKCU:\Software\Microsoft\Office\$v\Excel"
    if (Test-Path $probe) {
        $excelOptionsKey = "$probe\Options"
        Write-Host ("  检测到 Excel 版本：{0}" -f $v)
        break
    }
}
if (-not $excelOptionsKey) {
    throw "没有在注册表里找到 Excel，请确认这台电脑装的是 Windows 版 Microsoft Excel"
}

if (-not (Test-Path $excelOptionsKey)) {
    New-Item -Path $excelOptionsKey -Force | Out-Null
}

# 找一个没被其他加载项占用的 OPEN / OPENn 键名
$existingNames = (Get-Item $excelOptionsKey).GetValueNames()

$keyName = $null
if ($existingNames -notcontains "OPEN") {
    $keyName = "OPEN"
} else {
    $i = 1
    while ($existingNames -contains "OPEN$i") { $i++ }
    $keyName = "OPEN$i"
}

Set-ItemProperty -Path $excelOptionsKey -Name $keyName -Value ('"' + $xlamTarget + '"')

Write-Host ("  已写入注册表：{0}\{1}" -f $excelOptionsKey, $keyName)
Write-Host ""
Write-Host "安装完成！" -ForegroundColor Green
Write-Host "请完全关闭并重新打开 Excel（不只是关一个窗口，要确认后台进程也退出了），"
Write-Host "功能区应该会自动出现「韵词字幕工具」选项卡。"

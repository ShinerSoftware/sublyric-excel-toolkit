# PowerShell 安装脚本使用指南

`install/` 目录下的四个脚本，用来把打包好的加载项"装一次、以后不用管"，
不用每次都手动走"文件→选项→加载项"或者复制文件到 STARTUP 目录。

## 目录结构要求

脚本要求打包好的加载项文件和脚本放在**同一个目录**下：

**Excel：**
```
你的目录/
├── install-excel.ps1
├── uninstall-excel.ps1
├── ExcelToSRT.xlam       ← 按《01-Excel打包指南》做好的文件
└── reader.html           ← excel-addin/assets/reader.html
```

**Word：**
```
你的目录/
├── install-word.ps1
├── uninstall-word.ps1
├── SublyricWordTools.dotm  ← 按《02-Word打包指南》做好的文件
└── reader.html             ← 和 Excel 版共用同一个文件
```

（两边如果想放一起，脚本只认文件名，放同一个文件夹也没问题，两边的
`reader.html` 是完全一样的内容，放一份即可）

## 怎么跑

**方法一（推荐，最简单）**：右键脚本文件 → **"使用 PowerShell 运行"**

**方法二（命令行）**：
```powershell
cd 你的目录
powershell -ExecutionPolicy Bypass -File .\install-excel.ps1
```

如果双击/右键运行时 Windows 弹出"无法加载，因为在此系统上禁止运行脚本"，
说明这台电脑的 PowerShell 执行策略比较严格，用方法二的命令行方式绕过去即可
（`-ExecutionPolicy Bypass` 只影响这一次运行，不会永久改系统设置）。

## 装的是什么、卸载怎么卸

**Excel 版**（`install-excel.ps1`）：
1. 把 `ExcelToSRT.xlam` 和 `reader.html` 复制到 `%APPDATA%\Microsoft\AddIns`
2. 在当前 Windows 账户的注册表里加一条"开机自动加载"记录（和你手动在
   "加载项"里勾选是同一个效果，只是脚本代劳）

卸载用 `uninstall-excel.ps1`，会把上面两步都撤销（删文件 + 删注册表记录）。

**Word 版**（`install-word.ps1`）：
1. 把 `SublyricWordTools.dotm` 和 `reader.html` 复制到
   `%APPDATA%\Microsoft\Word\STARTUP`

Word 会自动加载 STARTUP 目录里的所有模板，不需要额外的注册表操作，
所以 Word 版脚本比 Excel 版简单。卸载用 `uninstall-word.ps1`，删掉这两个
文件即可。

## 安全性说明

- 全程只操作**当前 Windows 账户**的文件和注册表（`HKEY_CURRENT_USER`），
  不需要管理员权限，不会影响电脑上其他账户，也不会动系统级设置
- 脚本内容是纯文本，打开就能看到具体做了什么操作，不放心可以先通读一遍
  再运行
- 卸载脚本会完全撤销安装脚本做的事，装卸干净，不会有残留

## 常见问题

- **报错"找不到 xxx.xlam/dotm"**：确认加载项文件和脚本真的放在同一个目录，
  文件名要完全一致（区分大小写不敏感，但不能有多余的空格或后缀差异）
- **装完 Excel/Word 没出现新选项卡**：重新完全关闭 Excel/Word 再打开
  （不是关一个窗口，是确认后台进程也退出了），有些情况下 Office 要完全
  重启一次才会重新扫描加载项
- **杀毒软件/公司安全软件拦截脚本运行**：这种脚本本质上是文件复制 +
  注册表读写，如果公司电脑管得比较严，可能会被拦，需要找 IT 加白名单，
  或者退回手动安装的方式

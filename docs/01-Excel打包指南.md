# Excel 加载项打包指南

把 `excel-addin/` 目录下的源码，打包成一个能装进 Excel 的 `ExcelToSRT.xlam`。

---

## 第一步：新建加载项工作簿

1. 打开 Excel，新建一个空白工作簿
2. `文件` → `另存为`，文件类型选择 **Excel 加载宏 (*.xlam)**，文件名 `ExcelToSRT`
3. 保存后关闭，准备下一步重新打开它编辑

## 第二步：导入 VBA 模块

1. 重新打开 `ExcelToSRT.xlam`
2. `Alt + F11` 打开 VBA 编辑器
3. 左侧工程资源管理器右键该工程 → `导入文件`，依次导入 `excel-addin/src-gbk/` 目录下的三个文件：
   - `modCore.bas`
   - `modReader.bas`
   - `modRibbon.bas`

   > 为什么用 `src-gbk` 而不是 `src`：中文版 Excel 的 VBA 编辑器导入 `.bas` 文件时
   > 按系统默认编码（GBK）读取，如果导入 UTF-8 编码的文件，里面的中文注释和
   > 提示文字会变成乱码。`src/` 目录下是 UTF-8 源码（方便在电脑上直接阅读、
   > 版本管理），`src-gbk/` 是专门转码过、给 VBA 导入用的版本，内容完全一致。

4. `Ctrl + S` 保存

## 第三步：插入功能区图标

1. 下载安装 **Office RibbonX Editor**（免费开源）：
   https://github.com/fernandreu/office-ribbonx-editor/releases/latest
   下载文件名带 **SelfContained** 的版本最省事，不用额外装 .NET 运行时
2. 关闭 Excel（RibbonX Editor 打开文件时不能被 Excel 占用）
3. 打开 RibbonX Editor → `File` → `Open` → 选中 `ExcelToSRT.xlam`
4. 左侧树里找到当前这个文件 → 右键 → **"Insert Icon(s)"**
5. 一次性选中 `excel-addin/customUI/icons/` 目录下这 7 张图（`icon_extra_sync.png`
   是备用的，不用插）：

   | 图标文件 | 对应按钮 |
   |---|---|
   | `icon_format_sheet.png` | 基本格式处理 |
   | `icon_remove_symbols.png` | 删除句末符号 |
   | `icon_add_spaces.png` | 英文加空格 |
   | `icon_reader.png` | 朗读校对 |
   | `icon_export_bilingual.png` | 导出双语字幕 |
   | `icon_export_chinese.png` | 导出中文字幕 |
   | `icon_export_word.png` | 导出双语字幕文档 |

6. 插入后确认每张图标的名字（不带后缀）和文件名一致——RibbonX Editor 默认就是
   拿文件名当 ID，正常不用手动改

## 第四步：插入 Ribbon 定义

1. 左侧切到 **"Office 2010+ CustomUI Part"**
2. 把 `excel-addin/customUI/customUI14.xml` 的内容整个复制粘贴进去，覆盖默认模板
   （这份 XML 里的按钮已经用 `image="icon_xxx"` 引用了上一步插入的图标，
   不需要再改）
3. 工具栏点 **"Validate"**，确认没有报错
4. `File` → `Save`

## 第五步：放置 reader.html（朗读校对功能需要）

把 `excel-addin/assets/reader.html` 复制到和 `ExcelToSRT.xlam` **完全相同的目录**下：

```
你的安装目录/
├── ExcelToSRT.xlam
└── reader.html
```

代码里是按 `ThisWorkbook.Path & "\reader.html"` 找这个文件的，路径对不上会弹提示。

## 第六步：安装测试

手动测试可以直接：`文件` → `选项` → `加载项` → 底部选 `Excel 加载宏` → `转到`
→ `浏览` 选中 `.xlam` → 勾选启用。

如果要给别人用、或者要"装一次以后不用管"，参考 `install/` 目录下的 PowerShell
安装脚本，见 `docs/PowerShell使用指南.md`。

## 常见问题

- **看不到新选项卡**：确认加载项已勾选启用；也可能是 XML 有语法错误，回到
  RibbonX Editor 点 "Validate" 检查
- **按钮点击没反应/报错找不到过程**：检查 `customUI14.xml` 里每个按钮的
  `onAction` 名称，是否和 `modRibbon.bas` 里的 `Public Sub` 名称完全一致
- **图标显示不出来/显示成空白**：确认插入图标这一步，图标名字和 XML 里
  `image="..."` 引用的名字完全一致（区分大小写）

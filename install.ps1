# OfficeAutoSave 安装脚本（免管理员，仅写当前用户注册表 HKCU）
# 用法：右键本文件 → "使用 PowerShell 运行"；若被策略拦截，在本目录执行：
#   powershell -ExecutionPolicy Bypass -File .\install.ps1
$ErrorActionPreference = "Stop"

$src = Split-Path -Parent $MyInvocation.MyCommand.Path

# 把 DLL 复制到固定位置再注册（注册后该位置不能删/移动）
$dest = Join-Path $env:LOCALAPPDATA "OfficeAutoSave"
New-Item -ItemType Directory -Path $dest -Force | Out-Null
Copy-Item (Join-Path $src "*.dll") $dest -Force
$dll = Join-Path $dest "OfficeAutoSave.dll"
if (!(Test-Path $dll)) { throw "未找到 OfficeAutoSave.dll，请确认本脚本与编译产物在同一目录。" }

$guid = "{3F8A9C2E-7B4D-4E1F-9A2B-6C5D8E0F1A2B}"
$clsid = "HKCU:\Software\Classes\CLSID\$guid"
$inproc = "$clsid\InprocServer32"
$codebase = ([uri]$dll).AbsoluteUri

# 1) 注册托管 COM 类（mscoree 承载，CLR4）
New-Item -Path $inproc -Force | Out-Null
New-ItemProperty -Path $clsid -Name "(default)" -Value "OfficeAutoSave.Connect" -Force | Out-Null
New-ItemProperty -Path $inproc -Name "(default)" -Value "$env:SystemRoot\System32\mscoree.dll" -Force | Out-Null
New-ItemProperty -Path $inproc -Name "ThreadingModel" -Value "Both" -Force | Out-Null
New-ItemProperty -Path $inproc -Name "Class" -Value "OfficeAutoSave.Connect" -Force | Out-Null
New-ItemProperty -Path $inproc -Name "Assembly" -Value "OfficeAutoSave, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" -Force | Out-Null
New-ItemProperty -Path $inproc -Name "CodeBase" -Value $codebase -Force | Out-Null
New-ItemProperty -Path $inproc -Name "RuntimeVersion" -Value "v4.0.30319" -Force | Out-Null

# 2) 注册为 Word / Excel / PowerPoint 的加载项（LoadBehavior=3 表示随启动加载）
foreach ($app in "Word", "Excel", "PowerPoint") {
    $key = "HKCU:\Software\Microsoft\Office\$app\Addins\OfficeAutoSave.Connect"
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name "Description" -Value "定时自动保存已打开的文档" -Force | Out-Null
    New-ItemProperty -Path $key -Name "FriendlyName" -Value "OfficeAutoSave" -Force | Out-Null
    New-ItemProperty -Path $key -Name "LoadBehavior" -PropertyType DWord -Value 3 -Force | Out-Null
}

Write-Host ""
Write-Host "安装完成！请完全退出 Word/Excel/PowerPoint 后重新打开，功能区将出现 [自动保存] 选项卡。" -ForegroundColor Green
Write-Host "DLL 已复制到: $dest （请勿删除或移动该目录）"

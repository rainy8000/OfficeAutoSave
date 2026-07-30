# OfficeAutoSave 卸载脚本
$ErrorActionPreference = "Continue"

$guid = "{3F8A9C2E-7B4D-4E1F-9A2B-6C5D8E0F1A2B}"
Remove-Item "HKCU:\Software\Classes\CLSID\$guid" -Recurse -Force -ErrorAction SilentlyContinue

foreach ($app in "Word", "Excel", "PowerPoint") {
    Remove-Item "HKCU:\Software\Microsoft\Office\$app\Addins\OfficeAutoSave.Connect" -Recurse -Force -ErrorAction SilentlyContinue
}

Remove-Item (Join-Path $env:LOCALAPPDATA "OfficeAutoSave") -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "已卸载。完全退出并重新打开 Office 后生效。"

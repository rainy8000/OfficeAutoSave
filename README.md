# OfficeAutoSave

Office（Word / Excel / PowerPoint）定时自动保存插件：每 N 分钟（默认 5，可在 1–120 之间调整）
自动保存所有**已打开且有改动**的文档。纯本地运行，与宏无关，不受宏安全策略影响。

## 特性

- 一个 DLL 同时支持 Word / Excel / PowerPoint
- 功能区新增"自动保存"选项卡 → 设置窗口：开关、间隔、上次保存时间
- 只保存"有改动"的文档；跳过从未保存的新文档（避免反复弹"另存为"）
- 安装免管理员（仅写当前用户注册表），卸载一键还原
- 构建零门槛：不引用 Office PIA / VSTO，GitHub Actions 直接 `dotnet build`

## 安装（编译产物 zip 解压后）

1. 解压到一个固定目录（装好后目录内的 DLL 会被复制到 `%LOCALAPPDATA%\OfficeAutoSave`）
2. 右键 `install.ps1` → **使用 PowerShell 运行**
   - 若提示禁止运行脚本，在该目录打开 PowerShell 执行：
     `powershell -ExecutionPolicy Bypass -File .\install.ps1`
3. 完全退出 Word/Excel/PowerPoint 再打开 → 功能区出现 **"自动保存"** 选项卡

## 使用

- 选项卡 → **自动保存设置…**：启用/禁用、保存间隔（分钟）、查看上次自动保存时间
- 注意：**从未手动保存过的新文档不会被自动保存**（设计如此，避免每轮弹"另存为"），
  新文档请先手动保存一次，之后由插件接管
- 卸载：运行 `uninstall.ps1`，重启 Office

## 常见问题

- **看不到"自动保存"选项卡**：文件 → 选项 → 加载项 → 底部"管理：COM 加载项"→ 转到 →
  勾选 OfficeAutoSave；若出现在"禁用项目"里，先启用它
- **公司电脑**：本插件不写 HKLM、不需要管理员；但如果 IT 策略整体禁用了 COM 加载项，
  则任何插件方案都无法绕过
- 保存大文档瞬间有轻微卡顿，属正常现象

## 从源码构建

任何装了 .NET SDK 的 Windows 机器（无需安装 Office / Visual Studio）：

```
dotnet build src/OfficeAutoSave.csproj -c Release -o build-output
```

GitHub Actions 已配置好（`.github/workflows/build.yml`）：push 后自动编译，
在运行记录的 **Artifacts → OfficeAutoSave** 下载产物 zip。

## 技术说明

- 形态：托管 COM 加载项（`IDTExtensibility2` + `IRibbonExtensibility`），
  接口用 `ComImport` 自声明（GUID/DispId 与官方一致），宿主对象模型用 `dynamic` 晚期绑定
- 注册：HKCU 下的 CLSID + `Software\Microsoft\Office\{Word,Excel,PowerPoint}\Addins`，
  由 mscoree 承载 CLR4（.NET Framework 4.8，Win10/11 自带）
- 定时器：`System.Windows.Forms.Timer`（在 Office 主线程触发，避免跨线程 COM 调用）

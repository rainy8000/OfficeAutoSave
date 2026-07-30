namespace OfficeAutoSave
{
    /// <summary>功能区定义。内联为常量，避免嵌入式资源的额外配置。</summary>
    internal static class RibbonXml
    {
        internal const string Text = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabAutoSave"" label=""自动保存"">
        <group id=""grpAutoSave"" label=""自动保存"">
          <button id=""btnSettings"" label=""自动保存设置…""
                  size=""large"" onAction=""OnSettingsClick""
                  imageMso=""FileSave"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }
}

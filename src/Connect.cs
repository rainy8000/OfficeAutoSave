using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficeAutoSave
{
    /// <summary>
    /// 加载项入口。同时实现 IDTExtensibility2（Office 加载协议）和
    /// IRibbonExtensibility（功能区扩展），Office 启动时自动创建本类实例。
    /// 对宿主 Application 使用 dynamic 晚期绑定（Word/Excel/PowerPoint 通用），
    /// 因此编译期不引用任何 Office 互操作程序集。
    /// 注册信息由 install.ps1 写入 HKCU，安装免管理员。
    /// </summary>
    [ComVisible(true)]
    [Guid("3F8A9C2E-7B4D-4E1F-9A2B-6C5D8E0F1A2B")]
    [ProgId("OfficeAutoSave.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)] // 让 Office 能按名字回调 OnSettingsClick
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private dynamic _app;
        private Timer _timer;

        internal static DateTime? LastSaveTime { get; private set; }
        internal static int LastSavedCount { get; private set; }

        #region IDTExtensibility2 —— Office 加载/卸载生命周期

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _app = Application;
            RestartTimer();
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            StopTimer();
            _app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { StopTimer(); }

        #endregion

        #region IRibbonExtensibility —— 功能区

        public string GetCustomUI(string ribbonID)
        {
            return RibbonXml.Text;
        }

        // 功能区按钮回调（Office 通过 IDispatch 按名字调用本方法）
        public void OnSettingsClick(IRibbonControl control)
        {
            using (var form = new SettingsForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                    RestartTimer();
            }
        }

        #endregion

        #region 定时保存

        internal void RestartTimer()
        {
            StopTimer();
            if (!Config.Current.AutoSaveEnabled) return;
            int minutes = Math.Max(1, Config.Current.IntervalMinutes);
            _timer = new Timer { Interval = minutes * 60 * 1000 };
            _timer.Tick += OnTick;
            _timer.Start();
        }

        private void StopTimer()
        {
            if (_timer == null) return;
            _timer.Stop();
            _timer.Tick -= OnTick;
            _timer.Dispose();
            _timer = null;
        }

        private void OnTick(object sender, EventArgs e)
        {
            if (_app == null) return;
            int saved = 0;
            try
            {
                dynamic docs = GetOpenDocuments(_app);
                if (docs != null)
                {
                    foreach (dynamic doc in docs)
                    {
                        try
                        {
                            string path = doc.Path;
                            // 只保存：已落过盘（Path 非空）且有未保存改动
                            if (!string.IsNullOrEmpty(path) && !IsSaved(doc))
                            {
                                doc.Save();
                                saved++;
                            }
                        }
                        catch { /* 单个文档失败（只读/弹窗等）跳过 */ }
                    }
                }
            }
            catch { /* 宿主正显示模态对话框时访问 COM 会抛错，跳过本轮 */ }

            if (saved > 0)
            {
                LastSaveTime = DateTime.Now;
                LastSavedCount = saved;
            }
        }

        // Word=Documents，Excel=Workbooks，PowerPoint=Presentations
        private static dynamic GetOpenDocuments(dynamic app)
        {
            try { return app.Documents; } catch { }
            try { return app.Workbooks; } catch { }
            try { return app.Presentations; } catch { }
            return null;
        }

        // Word/Excel 的 Saved 是 bool；PowerPoint 是三态枚举（msoTrue=-1，msoFalse=0）。
        // 统一转成 "非 0 即已保存" 来兼容三种宿主。
        private static bool IsSaved(dynamic doc)
        {
            try
            {
                object v = doc.Saved;
                if (v is bool b) return b;
                return Convert.ToInt32(v) != 0;
            }
            catch
            {
                return true; // 拿不到状态就当已保存，不做多余动作
            }
        }

        #endregion
    }
}

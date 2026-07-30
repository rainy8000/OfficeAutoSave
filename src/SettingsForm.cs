using System;
using System.Drawing;
using System.Windows.Forms;

namespace OfficeAutoSave
{
    /// <summary>
    /// 自动保存设置窗。界面全部用代码构建，无需设计器文件。
    /// </summary>
    public class SettingsForm : Form
    {
        private readonly NumericUpDown _num;
        private readonly CheckBox _chk;

        public SettingsForm()
        {
            Text = "自动保存设置";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(330, 190);

            _chk = new CheckBox
            {
                Text = "启用自动保存",
                Left = 20,
                Top = 18,
                Width = 280,
                Checked = Config.Current.AutoSaveEnabled
            };

            var lbl = new Label
            {
                Text = "保存间隔（分钟）：",
                Left = 20,
                Top = 56,
                Width = 130
            };

            _num = new NumericUpDown
            {
                Left = 155,
                Top = 53,
                Width = 70,
                Minimum = 1,
                Maximum = 120,
                Value = Math.Min(120, Math.Max(1, Config.Current.IntervalMinutes))
            };

            var lblLast = new Label
            {
                Left = 20,
                Top = 92,
                Width = 290,
                Text = Connect.LastSaveTime.HasValue
                    ? "上次自动保存：" + Connect.LastSaveTime.Value.ToString("HH:mm:ss")
                      + "（" + Connect.LastSavedCount + " 个文档）"
                    : "本次运行尚未执行自动保存"
            };

            var btnOk = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Left = 145,
                Top = 138,
                Width = 75
            };
            btnOk.Click += (s, e) =>
            {
                Config.Current.AutoSaveEnabled = _chk.Checked;
                Config.Current.IntervalMinutes = (int)_num.Value;
                Config.Current.Save();
            };

            var btnCancel = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Left = 230,
                Top = 138,
                Width = 75
            };

            Controls.AddRange(new Control[] { _chk, lbl, _num, lblLast, btnOk, btnCancel });
            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }
    }
}

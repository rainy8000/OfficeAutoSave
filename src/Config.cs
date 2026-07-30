using System.Configuration;

namespace OfficeAutoSave
{
    /// <summary>
    /// 用户级设置（开关 + 间隔），自动持久化到 user.config，无需任何配置文件。
    /// </summary>
    internal class Config : ApplicationSettingsBase
    {
        private static readonly Config _current = new Config();
        internal static Config Current => _current;

        [UserScopedSetting]
        [DefaultSettingValue("true")]
        public bool AutoSaveEnabled
        {
            get => (bool)this[nameof(AutoSaveEnabled)];
            set => this[nameof(AutoSaveEnabled)] = value;
        }

        [UserScopedSetting]
        [DefaultSettingValue("5")]
        public int IntervalMinutes
        {
            get => (int)this[nameof(IntervalMinutes)];
            set => this[nameof(IntervalMinutes)] = value;
        }
    }
}

#pragma warning disable VSEXTPREVIEW_SETTINGS

using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.VisualStudio.Extensibility;
using Microsoft.VisualStudio.Extensibility.Settings;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json;

namespace TabsManagerExtension.Configuration {
    internal static class TabsManagerConfigurationService {
        private static readonly object _sync = new object();
        private static readonly string _configurationDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TabsManagerExtension");
        private static readonly string _configurationPath = Path.Combine(_configurationDirectory, "configuration.json");
        private static readonly List<IDisposable> _settingsSubscriptions = new List<IDisposable>();
        private static TabsManagerConfiguration? _configuration;
        private static VisualStudioExtensibility? _extensibility;
        private static bool _autoLoadCustomTabs = true;
        private static double _tabsScaleFactor = 1.0;
        private static bool _settingsInitialized;

        public static event Action? SettingsInitialized;
        public static event Action<double>? TabsScaleFactorChanged;

        public static bool IsSettingsInitialized {
            get {
                lock (_sync) {
                    return _settingsInitialized;
                }
            }
        }

        public static bool AutoLoadCustomTabs {
            get {
                lock (_sync) {
                    return _autoLoadCustomTabs;
                }
            }
        }

        public static double TabsScaleFactor {
            get {
                lock (_sync) {
                    return _tabsScaleFactor;
                }
            }
        }

        public static IReadOnlyList<string> OpenToolWindowIds {
            get {
                lock (_sync) {
                    return Load().OpenToolWindowIds.ToArray();
                }
            }
        }

        public static async Task InitializeAsync(VisualStudioExtensibility extensibility, CancellationToken cancellationToken) {
            lock (_sync) {
                if (_settingsInitialized) {
                    return;
                }

                _extensibility = extensibility;
            }

            var autoLoadResult = await extensibility.Settings().ReadEffectiveValueAsync(
                TabsManagerSettingDefinitions.AutoLoadCustomTabs,
                cancellationToken
            );

            var scaleResult = await extensibility.Settings().ReadEffectiveValueAsync(
                TabsManagerSettingDefinitions.TabsScaleFactor,
                cancellationToken
            );

            lock (_sync) {
                _autoLoadCustomTabs = autoLoadResult.ValueOrDefault(defaultValue: true);
                _tabsScaleFactor = NormalizeTabsScaleFactor((double)scaleResult.ValueOrDefault(defaultValue: 1.0m));
            }

            _settingsSubscriptions.Add(await extensibility.Settings().SubscribeAsync(
                TabsManagerSettingDefinitions.AutoLoadCustomTabs,
                cancellationToken,
                result => SetAutoLoadCustomTabsCache(result.ValueOrDefault(defaultValue: true))
            ));

            _settingsSubscriptions.Add(await extensibility.Settings().SubscribeAsync(
                TabsManagerSettingDefinitions.TabsScaleFactor,
                cancellationToken,
                result => SetTabsScaleFactorCache((double)result.ValueOrDefault(defaultValue: 1.0m))
            ));

            lock (_sync) {
                _settingsInitialized = true;
            }

            Helpers.Diagnostic.Logger.LogDebug($"[Settings] Настройки загружены: AutoLoadCustomTabs={_autoLoadCustomTabs}, TabsScaleFactor={_tabsScaleFactor}.");
            TabsScaleFactorChanged?.Invoke(_tabsScaleFactor);
            SettingsInitialized?.Invoke();
        }

        public static void SetAutoLoadCustomTabs(bool value) {
            lock (_sync) {
                if (_autoLoadCustomTabs == value) {
                    return;
                }

                _autoLoadCustomTabs = value;
            }

            WriteSettingAsync(
                batch => batch.WriteSetting(TabsManagerSettingDefinitions.AutoLoadCustomTabs, value),
                "Updating Tabs Manager automatic activation"
            );
        }

        public static void SetOpenToolWindowIds(IEnumerable<string> windowIds) {
            lock (_sync) {
                var configuration = Load();

                // В конфигурацию попадают только GUID окон без повторов, но в исходном порядке.
                var newWindowIds = windowIds.Where(id => Guid.TryParse(id, out _)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                if (configuration.OpenToolWindowIds.SequenceEqual(newWindowIds, StringComparer.OrdinalIgnoreCase)) {
                    return;
                }

                configuration.OpenToolWindowIds = newWindowIds;
                Save(configuration);
            }
        }

        public static void SetTabsScaleFactor(double value) {
            double normalizedValue = NormalizeTabsScaleFactor(value);
            lock (_sync) {
                if (Math.Abs(_tabsScaleFactor - normalizedValue) <= 0.001) {
                    return;
                }

                _tabsScaleFactor = normalizedValue;
            }

            WriteSettingAsync(
                batch => batch.WriteSetting(TabsManagerSettingDefinitions.TabsScaleFactor, (decimal)normalizedValue),
                "Updating Tabs Manager tab compactness"
            );
        }

        private static void SetTabsScaleFactorCache(double value) {
            double normalizedValue = NormalizeTabsScaleFactor(value);
            lock (_sync) {
                if (Math.Abs(_tabsScaleFactor - normalizedValue) <= 0.001) {
                    return;
                }

                _tabsScaleFactor = normalizedValue;
            }

            TabsScaleFactorChanged?.Invoke(normalizedValue);
        }

        private static void SetAutoLoadCustomTabsCache(bool value) {
            lock (_sync) {
                _autoLoadCustomTabs = value;
            }
        }

        private static void WriteSettingAsync(Action<SettingsWriteBatch> writeAction, string description) {
            var extensibility = _extensibility;
            if (extensibility == null) {
                return;
            }

            ThreadHelper.JoinableTaskFactory.RunAsync(async () => {
                try {
                    await extensibility.Settings().WriteAsync(writeAction, description, CancellationToken.None);
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogError($"Failed to write Tabs Manager setting: {ex}");
                }
            }).FileAndForget("TabsManagerExtension/WriteSetting");
        }

        private static double NormalizeTabsScaleFactor(double value) {
            return double.IsNaN(value) || double.IsInfinity(value)
                ? 1.0
                : Helpers.Math.Clamp(value, 0.5, 1.5);
        }

        private static TabsManagerConfiguration Load() {
            if (_configuration != null) {
                return _configuration;
            }

            try {
                if (File.Exists(_configurationPath)) {
                    _configuration = JsonConvert.DeserializeObject<TabsManagerConfiguration>(File.ReadAllText(_configurationPath));
                }
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to load Tabs Manager configuration: {ex}");
            }

            _configuration ??= new TabsManagerConfiguration();
            _configuration.OpenToolWindowIds ??= new List<string>();
            return _configuration;
        }

        private static void Save(TabsManagerConfiguration configuration) {
            try {
                Directory.CreateDirectory(_configurationDirectory);
                var temporaryPath = _configurationPath + ".tmp";
                File.WriteAllText(temporaryPath, JsonConvert.SerializeObject(configuration, Formatting.Indented));

                if (File.Exists(_configurationPath)) {
                    File.Replace(temporaryPath, _configurationPath, null);
                }
                else {
                    File.Move(temporaryPath, _configurationPath);
                }
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to save Tabs Manager configuration: {ex}");
            }
        }
    }
}

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json;


namespace TabsManagerExtension.Configuration {
    internal static class TabsManagerConfigurationService {
        private static readonly object _sync = new object();
        private static readonly string _configurationDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TabsManagerExtension");
        private static readonly string _configurationPath = Path.Combine(_configurationDirectory, "configuration.json");
        private static TabsManagerConfiguration? _configuration;

        public static bool AutoLoadCustomTabs {
            get {
                lock (_sync) {
                    return Load().AutoLoadCustomTabs;
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

        public static void SetAutoLoadCustomTabs(bool value) {
            lock (_sync) {
                var configuration = Load();
                configuration.AutoLoadCustomTabs = value;
                Save(configuration);
            }
        }

        public static void SetOpenToolWindowIds(IEnumerable<string> windowIds) {
            lock (_sync) {
                var configuration = Load();

                // В конфигурацию попадают только GUID окон без повторов, но в исходном порядке.
                var newWindowIds = windowIds.Where(id => Guid.TryParse(id, out _)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

                // Метод вызывается из UI-событий: не перезаписываем JSON, если набор окон не изменился.
                if (configuration.OpenToolWindowIds.SequenceEqual(newWindowIds, StringComparer.OrdinalIgnoreCase)) {
                    return;
                }

                configuration.OpenToolWindowIds = newWindowIds;
                Save(configuration);
            }
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

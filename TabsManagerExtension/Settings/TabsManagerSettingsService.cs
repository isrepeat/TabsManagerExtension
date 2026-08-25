#pragma warning disable VSEXTPREVIEW_SETTINGS

using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Extensibility;
using Microsoft.VisualStudio.Extensibility.Settings;
using Newtonsoft.Json;

namespace TabsManagerExtension.Settings {
    // Единственный источник пользовательских значений — локальный settings.json. Unified Settings используется
    // только как мост для нескольких элементов стандартного окна Options и не хранит оформление панели.
    internal static class TabsManagerSettingsService {
        private sealed class LegacyToolWindowState {
            public List<string>? OpenToolWindowIds { get; set; }
            public string? ActiveToolWindowId { get; set; }
        }

        public const string DefaultAnchorSectionPattern = @"^\s*//\s*░(?!░)\s*(?<title>[^\r\n]+?)\s*\r?\n\s*//\s*░{3,}\s*(?=\r?\n|$)";
        public const string DefaultAnchorSubsectionPattern = @"^\s*//\s*░(?!░)\s*(?<title>[^\r\n]+?)\s*(?=\r?\n|$)";

        private static readonly object _sync = new object();
        private static readonly string _settingsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TabsManagerExtension");
        private static readonly string _settingsPath = Path.Combine(_settingsDirectory, "settings.json");
        private static readonly string _legacySettingsPath = Path.Combine(_settingsDirectory, "configuration.json");
        private static readonly List<IDisposable> _settingsSubscriptions = new List<IDisposable>();
        private static readonly CancellationTokenSource _shutdownCancellation = new CancellationTokenSource();
        private static TabsManagerSettings? _settings;
        private static VisualStudioExtensibility? _extensibility;
        private static bool _isShuttingDown;
        private static bool _autoLoadCustomTabs = true;
        private static double _tabsScaleFactor = 1.0;
        private static string _anchorSectionPattern = DefaultAnchorSectionPattern;
        private static string _anchorSubsectionPattern = DefaultAnchorSubsectionPattern;
        private static int _anchorPatternsRevision;
        private static readonly Dictionary<string, object> _appearance = new Dictionary<string, object>(StringComparer.Ordinal) {
            ["panelBackgroundColor"] = "#FF252526", ["tabBackgroundColor"] = "#00FFFFFF", ["tabBorderColor"] = "#00FFFFFF",
            ["tabHoverBackgroundColor"] = "#663C3C3C", ["tabHoverBorderColor"] = "#00FFFFFF",
            ["selectedTabBackgroundColor"] = "#663C3C3C", ["selectedTabBorderColor"] = "#00FFFFFF",
            ["activeTabBackgroundColor"] = "#FF2D2D30", ["activeTabBorderColor"] = "#FF4B4B4B",
            ["tabTextColor"] = "#FF808080", ["tabTextBold"] = false, ["tabTextItalic"] = false, ["tabTextSize"] = 12d,
            ["tabHoverTextColor"] = "#FFFFFFFF", ["tabHoverTextBold"] = false, ["tabHoverTextItalic"] = false, ["tabHoverTextSize"] = 12d,
            ["selectedTabTextColor"] = "#FFFFFFFF", ["selectedTabTextBold"] = true, ["selectedTabTextItalic"] = false, ["selectedTabTextSize"] = 12d,
            ["activeTabTextColor"] = "#FFFFFFFF", ["activeTabTextBold"] = true, ["activeTabTextItalic"] = false, ["activeTabTextSize"] = 12d
        };
        private static bool _settingsInitialized;

        static TabsManagerSettingsService() {
            // Синхронная загрузка нужна, чтобы оформление и режим вкладок были доступны до построения UI.
            LoadLocalSettings();
            _settingsInitialized = true;
        }

        public static event Action? SettingsInitialized;
        public static event Action<double>? TabsScaleFactorChanged;
        public static event Action? AnchorPatternsChanged;
        public static event Action? AppearanceChanged;
        public static event Action<bool>? AutoLoadCustomTabsChanged;
        public static event Action? ToolbarButtonsVisibilityChanged;

        public static string GetAppearanceColor(string key) {
            lock (_sync) {
                return (string)_appearance[key];
            }
        }

        public static bool GetAppearanceBoolean(string key) {
            lock (_sync) {
                return (bool)_appearance[key];
            }
        }

        public static double GetAppearanceNumber(string key) {
            lock (_sync) {
                return (double)_appearance[key];
            }
        }

        public static void SetAppearanceColor(string key, string value) {
            string defaultValue = GetAppearanceColor(key);
            SetAppearanceColor(key, value, defaultValue);
        }

        public static void SetAppearanceBoolean(string key, bool value) {
            SetAppearanceBoolean(key, value, notify: true);
        }

        public static void SetAppearanceNumber(string key, double value) {
            double normalizedValue = Helpers.Math.Clamp(value, 8, 32);
            SetAppearanceNumber(key, normalizedValue, notify: true);
        }

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

        public static string AnchorSectionPattern {
            get {
                lock (_sync) {
                    return _anchorSectionPattern;
                }
            }
        }

        public static string AnchorSubsectionPattern {
            get {
                lock (_sync) {
                    return _anchorSubsectionPattern;
                }
            }
        }

        public static int AnchorPatternsRevision {
            get {
                lock (_sync) {
                    return _anchorPatternsRevision;
                }
            }
        }

        public static string ActiveSettingsSection {
            get {
                lock (_sync) {
                    return Load().ActiveSettingsSection;
                }
            }
        }

        public static bool ShowTabsToggleToolbarButton => Load().ShowTabsToggleToolbarButton;
        public static bool ShowStandardTabsLayoutToolbarButton => Load().ShowStandardTabsLayoutToolbarButton;

        public static IReadOnlyList<string> OpenToolWindowIds {
            get {
                lock (_sync) {
                    return Load().OpenToolWindowIds.ToArray();
                }
            }
        }

        public static string? ActiveToolWindowId {
            get {
                lock (_sync) {
                    return Load().ActiveToolWindowId;
                }
            }
        }

        public static async Task InitializeAsync(VisualStudioExtensibility extensibility, CancellationToken cancellationToken) {
            // Асинхронная часть подключает стандартное окно Options после готовности Extensibility API.
            lock (_sync) {
                if (_extensibility != null || _isShuttingDown) {
                    return;
                }

                _extensibility = extensibility;
                _anchorPatternsRevision++;
            }

            using var initializationCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken,
                _shutdownCancellation.Token
            );
            var initializationToken = initializationCancellation.Token;

            Helpers.Diagnostic.Logger.LogDebug($"[Settings] Local settings.json loaded: AutoLoadCustomTabs={_autoLoadCustomTabs}, TabsScaleFactor={_tabsScaleFactor}.");
            TabsScaleFactorChanged?.Invoke(_tabsScaleFactor);
            AnchorPatternsChanged?.Invoke();
            AppearanceChanged?.Invoke();
            SettingsInitialized?.Invoke();

            try {
                await extensibility.Settings().WriteAsync(
                    batch => {
                        batch.WriteSetting(TabsManagerSettingsDefinitions.AutoLoadCustomTabs, _autoLoadCustomTabs);
                        batch.WriteSetting(TabsManagerSettingsDefinitions.ShowTabsToggleToolbarButton, ShowTabsToggleToolbarButton);
                        batch.WriteSetting(TabsManagerSettingsDefinitions.ShowStandardTabsLayoutToolbarButton, ShowStandardTabsLayoutToolbarButton);
                    },
                    "Synchronizing the Tabs Manager standard Options controls",
                    initializationToken
                );

                initializationToken.ThrowIfCancellationRequested();
                RegisterSettingsSubscription(await extensibility.Settings().SubscribeAsync(
                    TabsManagerSettingsDefinitions.AutoLoadCustomTabs,
                    initializationToken,
                    result => SetAutoLoadCustomTabs(result.ValueOrDefault(defaultValue: _autoLoadCustomTabs))
                ));

                initializationToken.ThrowIfCancellationRequested();
                RegisterSettingsSubscription(await extensibility.Settings().SubscribeAsync(
                    TabsManagerSettingsDefinitions.ShowTabsToggleToolbarButton,
                    initializationToken,
                    result => SetShowTabsToggleToolbarButton(result.ValueOrDefault(defaultValue: ShowTabsToggleToolbarButton))
                ));

                initializationToken.ThrowIfCancellationRequested();
                RegisterSettingsSubscription(await extensibility.Settings().SubscribeAsync(
                    TabsManagerSettingsDefinitions.ShowStandardTabsLayoutToolbarButton,
                    initializationToken,
                    result => SetShowStandardTabsLayoutToolbarButton(result.ValueOrDefault(defaultValue: ShowStandardTabsLayoutToolbarButton))
                ));

                initializationToken.ThrowIfCancellationRequested();
                // Unified Settings используется только как UI-мост для контролов стандартного Options.
                // Источником истины для пользовательских значений остаётся наш settings.json.
                RegisterSettingsSubscription(await extensibility.Settings().SubscribeAsync(
                    TabsManagerSettingsDefinitions.OpenSettingsPage,
                    initializationToken,
                    result => HandleOpenSettingsRequest(result.ValueOrDefault(defaultValue: false))
                ));
            }
            catch (Exception ex) when (IsExpectedSettingsShutdown(ex, initializationToken)) {
                Helpers.Diagnostic.Logger.LogDebug("[Settings] Unified Settings initialization stopped during Visual Studio shutdown.");
            }
        }

        public static void Shutdown() {
            IDisposable[] subscriptions;
            lock (_sync) {
                _isShuttingDown = true;
                _extensibility = null;
                subscriptions = _settingsSubscriptions.ToArray();
                _settingsSubscriptions.Clear();
            }

            _shutdownCancellation.Cancel();

            foreach (var subscription in subscriptions) {
                subscription.Dispose();
            }
        }

        private static void RegisterSettingsSubscription(IDisposable subscription) {
            lock (_sync) {
                if (!_isShuttingDown) {
                    _settingsSubscriptions.Add(subscription);
                    return;
                }
            }

            // Подписка могла завершить создание одновременно с выгрузкой пакета.
            subscription.Dispose();
        }

        private static bool IsExpectedSettingsShutdown(Exception exception, CancellationToken cancellationToken) {
            if (cancellationToken.IsCancellationRequested || exception is OperationCanceledException || exception is ObjectDisposedException) {
                return true;
            }

            return exception is AggregateException aggregateException &&
                aggregateException.InnerExceptions.All(innerException => IsExpectedSettingsShutdown(innerException, cancellationToken));
        }

        private static void HandleOpenSettingsRequest(bool openRequested) {
            if (!openRequested) {
                return;
            }

            ThreadHelper.JoinableTaskFactory.RunAsync(async () => {
                try {
                    await TabsManagerExtensionPackage.ShowCustomSettingsAsync();
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogError($"Failed to open Tabs Manager settings tab: {ex}");
                }
                finally {
                    WriteSettingAsync(
                        batch => batch.WriteSetting(TabsManagerSettingsDefinitions.OpenSettingsPage, false),
                        "Resetting the Tabs Manager settings launcher"
                    );
                }
            }).FileAndForget("TabsManagerExtension/OpenSettings");
        }

        public static void ResetAppearanceToDefaults() {
            SetAppearanceColor("panelBackgroundColor", "#FF252526", "#FF252526", false);
            SetAppearanceColor("tabBackgroundColor", "#00FFFFFF", "#00FFFFFF", false);
            SetAppearanceColor("tabBorderColor", "#00FFFFFF", "#00FFFFFF", false);
            SetAppearanceColor("tabHoverBackgroundColor", "#663C3C3C", "#663C3C3C", false);
            SetAppearanceColor("tabHoverBorderColor", "#00FFFFFF", "#00FFFFFF", false);
            SetAppearanceColor("selectedTabBackgroundColor", "#663C3C3C", "#663C3C3C", false);
            SetAppearanceColor("selectedTabBorderColor", "#00FFFFFF", "#00FFFFFF", false);
            SetAppearanceColor("activeTabBackgroundColor", "#FF2D2D30", "#FF2D2D30", false);
            SetAppearanceColor("activeTabBorderColor", "#FF4B4B4B", "#FF4B4B4B", false);
            SetAppearanceColor("tabTextColor", "#FF808080", "#FF808080", false);
            SetAppearanceBoolean("tabTextBold", false, false);
            SetAppearanceBoolean("tabTextItalic", false, false);
            SetAppearanceNumber("tabTextSize", 12, false);
            SetAppearanceColor("tabHoverTextColor", "#FFFFFFFF", "#FFFFFFFF", false);
            SetAppearanceBoolean("tabHoverTextBold", false, false);
            SetAppearanceBoolean("tabHoverTextItalic", false, false);
            SetAppearanceNumber("tabHoverTextSize", 12, false);
            SetAppearanceColor("selectedTabTextColor", "#FFFFFFFF", "#FFFFFFFF", false);
            SetAppearanceBoolean("selectedTabTextBold", true, false);
            SetAppearanceBoolean("selectedTabTextItalic", false, false);
            SetAppearanceNumber("selectedTabTextSize", 12, false);
            SetAppearanceColor("activeTabTextColor", "#FFFFFFFF", "#FFFFFFFF", false);
            SetAppearanceBoolean("activeTabTextBold", true, false);
            SetAppearanceBoolean("activeTabTextItalic", false, false);
            SetAppearanceNumber("activeTabTextSize", 12, false);
            AppearanceChanged?.Invoke();
            PersistLocalSettings();
        }

        public static void SetAutoLoadCustomTabs(bool value) {
            lock (_sync) {
                if (_autoLoadCustomTabs == value) {
                    return;
                }

                _autoLoadCustomTabs = value;
            }

            PersistLocalSettings();
            AutoLoadCustomTabsChanged?.Invoke(value);

            var extensibility = _extensibility;
            if (extensibility != null) {
                ThreadHelper.JoinableTaskFactory.RunAsync(async () => {
                    await extensibility.Settings().WriteAsync(
                        batch => batch.WriteSetting(TabsManagerSettingsDefinitions.AutoLoadCustomTabs, value),
                        "Synchronizing the Tabs Manager tabs switch",
                        CancellationToken.None
                    );
                }).FileAndForget("TabsManagerExtension/SyncTabsSwitch");
            }
        }

        public static void SetActiveSettingsSection(string section) {
            string normalizedSection = section == "customization" || section == "anchors" ? section : "main";
            lock (_sync) {
                var settings = Load();
                if (string.Equals(settings.ActiveSettingsSection, normalizedSection, StringComparison.Ordinal)) {
                    return;
                }

                settings.ActiveSettingsSection = normalizedSection;
                Save(settings);
            }
        }

        public static void SetShowTabsToggleToolbarButton(bool value) {
            SetToolbarButtonVisibility(
                value,
                settings => settings.ShowTabsToggleToolbarButton,
                (settings, visible) => settings.ShowTabsToggleToolbarButton = visible,
                TabsManagerSettingsDefinitions.ShowTabsToggleToolbarButton
            );
        }

        public static void SetShowStandardTabsLayoutToolbarButton(bool value) {
            SetToolbarButtonVisibility(
                value,
                settings => settings.ShowStandardTabsLayoutToolbarButton,
                (settings, visible) => settings.ShowStandardTabsLayoutToolbarButton = visible,
                TabsManagerSettingsDefinitions.ShowStandardTabsLayoutToolbarButton
            );
        }

        private static void SetToolbarButtonVisibility(
            bool value,
            Func<TabsManagerSettings, bool> getter,
            Action<TabsManagerSettings, bool> setter,
            Setting.Boolean unifiedSetting
        ) {
            lock (_sync) {
                var settings = Load();
                if (getter(settings) == value) {
                    return;
                }

                setter(settings, value);
                Save(settings);
            }

            ToolbarButtonsVisibilityChanged?.Invoke();
            var extensibility = _extensibility;
            if (extensibility != null) {
                ThreadHelper.JoinableTaskFactory.RunAsync(async () => {
                    await extensibility.Settings().WriteAsync(
                        batch => batch.WriteSetting(unifiedSetting, value),
                        "Synchronizing a Tabs Manager toolbar visibility switch",
                        CancellationToken.None
                    );
                }).FileAndForget("TabsManagerExtension/SyncToolbarVisibility");
            }
        }

        public static void SetAnchorSectionPattern(string value) {
            SetAnchorSectionPatternCache(value);
            PersistLocalSettings();
        }

        public static void SetAnchorSubsectionPattern(string value) {
            SetAnchorSubsectionPatternCache(value);
            PersistLocalSettings();
        }

        public static void SetOpenToolWindowState(IEnumerable<string> windowIds, string? activeWindowId) {
            lock (_sync) {
                var settings = Load();

                // В настройки попадают только GUID окон без повторов, но в исходном порядке.
                var newWindowIds = NormalizeToolWindowIds(windowIds);
                string? newActiveWindowId = NormalizeActiveToolWindowId(activeWindowId, newWindowIds);
                if (settings.OpenToolWindowIds.SequenceEqual(newWindowIds, StringComparer.OrdinalIgnoreCase) &&
                    string.Equals(settings.ActiveToolWindowId, newActiveWindowId, StringComparison.OrdinalIgnoreCase)) {

                    return;
                }

                settings.OpenToolWindowIds = newWindowIds;
                settings.ActiveToolWindowId = newActiveWindowId;
                Save(settings);
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

            PersistLocalSettings();
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

        private static void SetAnchorSectionPatternCache(string value) {
            string normalizedValue = NormalizeAnchorPattern(value, DefaultAnchorSectionPattern, "section");
            lock (_sync) {
                if (string.Equals(_anchorSectionPattern, normalizedValue, StringComparison.Ordinal)) {
                    return;
                }

                _anchorSectionPattern = normalizedValue;
                _anchorPatternsRevision++;
            }

            AnchorPatternsChanged?.Invoke();
        }

        private static void SetAnchorSubsectionPatternCache(string value) {
            string normalizedValue = NormalizeAnchorPattern(value, DefaultAnchorSubsectionPattern, "subsection");
            lock (_sync) {
                if (string.Equals(_anchorSubsectionPattern, normalizedValue, StringComparison.Ordinal)) {
                    return;
                }

                _anchorSubsectionPattern = normalizedValue;
                _anchorPatternsRevision++;
            }

            AnchorPatternsChanged?.Invoke();
        }

        private static string NormalizeAnchorPattern(string value, string defaultValue, string patternKind) {
            if (string.IsNullOrWhiteSpace(value)) {
                return defaultValue;
            }

            try {
                var regex = new Regex(value, RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
                if (!regex.GetGroupNames().Contains("title", StringComparer.Ordinal)) {
                    Helpers.Diagnostic.Logger.LogWarning($"[Settings] Anchor {patternKind} pattern has no named 'title' group. Default pattern is used.");
                    return defaultValue;
                }

                return value;
            }
            catch (ArgumentException ex) {
                Helpers.Diagnostic.Logger.LogWarning($"[Settings] Invalid anchor {patternKind} pattern. Default pattern is used: {ex.Message}");
                return defaultValue;
            }
        }

        private static void SetAppearanceColor(string key, string value, string defaultValue, bool notify = true) {
            string normalized = Regex.IsMatch(value ?? string.Empty, "^#(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$") ? value : defaultValue;
            SetAppearanceValue(key, normalized, notify);
        }

        private static void SetAppearanceBoolean(string key, bool value, bool notify = true) {
            SetAppearanceValue(key, value, notify);
        }

        private static void SetAppearanceNumber(string key, double value, bool notify = true) {
            SetAppearanceValue(key, Helpers.Math.Clamp(value, 8, 32), notify);
        }

        private static void SetAppearanceValue(string key, object value, bool notify) {
            lock (_sync) {
                if (Equals(_appearance[key], value)) {
                    return;
                }

                _appearance[key] = value;
            }

            if (notify) {
                PersistLocalSettings();
                AppearanceChanged?.Invoke();
            }
        }

        private static void LoadLocalSettings() {
            // Модель нормализует отсутствующие поля, затем типизированный кэш обслуживает горячие чтения из UI.
            var settings = Load();
            _autoLoadCustomTabs = settings.AutoLoadCustomTabs;
            _tabsScaleFactor = NormalizeTabsScaleFactor(settings.TabsScaleFactor);
            _anchorSectionPattern = NormalizeAnchorPattern(settings.AnchorSectionPattern, DefaultAnchorSectionPattern, "section");
            _anchorSubsectionPattern = NormalizeAnchorPattern(settings.AnchorSubsectionPattern, DefaultAnchorSubsectionPattern, "subsection");

            var cachedAppearance = settings.Appearance;
            foreach (var pair in cachedAppearance) {
                if (!_appearance.TryGetValue(pair.Key, out object currentValue)) {
                    continue;
                }

                if (currentValue is bool && bool.TryParse(pair.Value, out bool booleanValue)) {
                    _appearance[pair.Key] = booleanValue;
                }
                else if (currentValue is double && double.TryParse(pair.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double numberValue)) {
                    _appearance[pair.Key] = Helpers.Math.Clamp(numberValue, 8, 32);
                }
                else if (currentValue is string && Regex.IsMatch(pair.Value ?? string.Empty, "^#(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$")) {
                    _appearance[pair.Key] = pair.Value;
                }
            }
        }

        private static void PersistLocalSettings() {
            // Под блокировкой формируется согласованный снимок всех кэшированных значений.
            lock (_sync) {
                var settings = Load();
                settings.Version = 2;
                settings.AutoLoadCustomTabs = _autoLoadCustomTabs;
                settings.TabsScaleFactor = _tabsScaleFactor;
                settings.AnchorSectionPattern = _anchorSectionPattern;
                settings.AnchorSubsectionPattern = _anchorSubsectionPattern;
                settings.Appearance = _appearance.ToDictionary(
                    pair => pair.Key,
                    pair => Convert.ToString(pair.Value, CultureInfo.InvariantCulture) ?? string.Empty,
                    StringComparer.Ordinal
                );
                Save(settings);
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

        private static List<string> NormalizeToolWindowIds(IEnumerable<string>? windowIds) {
            return windowIds?
                .Where(id => Guid.TryParse(id, out _))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();
        }

        private static string? NormalizeActiveToolWindowId(string? activeWindowId, IReadOnlyCollection<string> windowIds) {
            return Guid.TryParse(activeWindowId, out _) && windowIds.Contains(activeWindowId!, StringComparer.OrdinalIgnoreCase)
                ? activeWindowId
                : null;
        }

        private static TabsManagerSettings Load() {
            if (_settings != null) {
                return _settings;
            }

            try {
                if (File.Exists(_settingsPath)) {
                    _settings = JsonConvert.DeserializeObject<TabsManagerSettings>(File.ReadAllText(_settingsPath));
                }
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to load Tabs Manager settings: {ex}");
            }

            _settings ??= new TabsManagerSettings();

            // configuration.json мог обновляться старой версией уже после создания settings.json.
            // Поэтому переносим из него только состояние tool window, не затрагивая новые настройки.
            bool migratedLegacySettings = TryMergeLegacyToolWindowState(_settings);
            _settings.AnchorSectionPattern ??= DefaultAnchorSectionPattern;
            _settings.AnchorSubsectionPattern ??= DefaultAnchorSubsectionPattern;
            _settings.ActiveSettingsSection ??= "main";
            _settings.OpenToolWindowIds = NormalizeToolWindowIds(_settings.OpenToolWindowIds);
            _settings.ActiveToolWindowId = NormalizeActiveToolWindowId(_settings.ActiveToolWindowId, _settings.OpenToolWindowIds);
            _settings.Appearance ??= new Dictionary<string, string>();

            if (migratedLegacySettings) {
                _settings.Version = 2;
                if (Save(_settings)) {
                    DeleteLegacySettings();
                }
            }

            return _settings;
        }

        private static bool TryMergeLegacyToolWindowState(TabsManagerSettings settings) {
            if (!File.Exists(_legacySettingsPath)) {
                return false;
            }

            try {
                var legacyState = JsonConvert.DeserializeObject<LegacyToolWindowState>(File.ReadAllText(_legacySettingsPath));
                if (legacyState == null) {
                    return false;
                }

                settings.OpenToolWindowIds = NormalizeToolWindowIds(legacyState.OpenToolWindowIds);
                settings.ActiveToolWindowId = NormalizeActiveToolWindowId(
                    legacyState.ActiveToolWindowId,
                    settings.OpenToolWindowIds
                );

                return true;
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to migrate legacy Tabs Manager configuration: {ex}");
                return false;
            }
        }

        private static void DeleteLegacySettings() {
            try {
                File.Delete(_legacySettingsPath);
                Helpers.Diagnostic.Logger.LogDebug("[Settings] Legacy configuration.json migrated to settings.json and removed.");
            }
            catch (Exception ex) {
                // При сбое удаления оставляем файл: следующий запуск безопасно повторит миграцию.
                Helpers.Diagnostic.Logger.LogWarning($"Failed to remove legacy Tabs Manager configuration: {ex}");
            }
        }

        private static bool Save(TabsManagerSettings settings) {
            try {
                Directory.CreateDirectory(_settingsDirectory);
                // Сначала пишем полный временный файл, чтобы сбой процесса не оставил частично записанный JSON.
                var temporaryPath = _settingsPath + ".tmp";
                File.WriteAllText(temporaryPath, JsonConvert.SerializeObject(settings, Formatting.Indented));

                if (File.Exists(_settingsPath)) {
                    File.Replace(temporaryPath, _settingsPath, null);
                }
                else {
                    File.Move(temporaryPath, _settingsPath);
                }

                return true;
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to save Tabs Manager settings: {ex}");
                return false;
            }
        }
    }
}

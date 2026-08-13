using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.ComponentModel.Design;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Utilities.UnifiedSettings;
using Microsoft.VisualStudio.VCProjectEngine;
using Microsoft.VisualStudio.VCCodeModel;
using Microsoft.VisualStudio.TextManager.Interop;
using Task = System.Threading.Tasks.Task;


namespace TabsManagerExtension {
    public class VsixVisualTreeHelper : Helpers.ObservableObject {
        private const string DocumentTabsLayoutSetting = "environment.tabs.documentTabs.layout";

        // Visual Studio хранит новые настройки интерфейса во внутренней подсистеме Unified Settings.
        // Получить её объект можно через Package.GetGlobalService(), передав идентификатор сервиса (SID).
        // В SDK 17.14 ещё нет готового класса с этим SID, поэтому объявляем локальный пустой класс
        // с нужным Guid. Экземпляры этого класса не создаются: его тип нужен только как ключ поиска.
        [Guid("E3684F31-344E-42EA-9047-B620FDC7AC25")]
        private sealed class SVsUnifiedSettingsManagerService {
        }

        private static readonly VsixVisualTreeHelper _instance = new();
        public static VsixVisualTreeHelper Instance => _instance;

        public bool IsCustomTabsInjected {
            get {
                return _currentTabHost?.TryGetTarget(out var decorator) == true &&
                       decorator.Child is Controls.TabsManagerToolWindowControl;
            }
        }

        private bool _isCustomTabsEnabled = false;
        public bool IsCustomTabsEnabled {
            get => _isCustomTabsEnabled;
            private set {
                if (_isCustomTabsEnabled != value) {
                    _isCustomTabsEnabled = value;
                    OnPropertyChanged();
                }
            }
        }


        private UIElement? _originalTabListHostContent;
        private WeakReference<Decorator>? _currentTabHost;
        private WeakReference<UIElement>? _lastInjectedContent;
        private string? _documentTabsLayoutBeforeCustomTabs;

        private VsixVisualTreeHelper() {
        }

        private static ISettingsWriter? GetSettingsWriter() {
            ThreadHelper.ThrowIfNotOnUIThread();

            Helpers.Diagnostic.Logger.LogDebug(
                $"Requesting Unified Settings service by SID '{typeof(SVsUnifiedSettingsManagerService).GUID}'."
            );

            var settingsService = Package.GetGlobalService(typeof(SVsUnifiedSettingsManagerService));

            Helpers.Diagnostic.Logger.LogDebug(
                $"Unified Settings service result: '{settingsService?.GetType().AssemblyQualifiedName ?? "<null>"}'."
            );

            var settingsManager = settingsService as ISettingsManager;
            if (settingsManager == null) {
                Helpers.Diagnostic.Logger.LogWarning("Unified settings manager is unavailable.");
                return null;
            }

            var settingsWriter = settingsManager.GetWriter("Tabs Manager");

            Helpers.Diagnostic.Logger.LogDebug(
                $"Unified Settings writer result: '{settingsWriter?.GetType().AssemblyQualifiedName ?? "<null>"}'."
            );

            return settingsWriter;
        }

        /// <summary>
        /// PART_TabListHost — стандартная область Visual Studio, в которой отображаются вкладки документов.
        /// Наш контрол подменяет содержимое этой области. Перед подменой переносим её влево,
        /// потому что интерфейс Tabs Manager рассчитан на вертикальную панель.
        /// </summary>
        public void EnsureStandardDocumentTabsOnLeft() {
            ThreadHelper.ThrowIfNotOnUIThread();

            this.SetDocumentTabsLayout("left", "Tabs Manager requires document tabs on the left");
        }

        public string? ToggleStandardDocumentTabsLayout() {
            ThreadHelper.ThrowIfNotOnUIThread();

            Helpers.Diagnostic.Logger.LogDebug(
                $"Toggle standard tabs layout requested; custom tabs enabled = {this.IsCustomTabsEnabled}."
            );

            if (this.IsCustomTabsEnabled) {
                Helpers.Diagnostic.Logger.LogWarning(
                    "Disable Tabs Manager before changing the standard document tabs layout."
                );

                return null;
            }

            var currentLayout = this.GetDocumentTabsLayout();
            if (string.IsNullOrWhiteSpace(currentLayout)) {
                return null;
            }

            var targetLayout = string.Equals(currentLayout, "left", StringComparison.OrdinalIgnoreCase)
                ? "top"
                : "left";

            Helpers.Diagnostic.Logger.LogDebug(
                $"Switching standard tabs layout from '{currentLayout}' to '{targetLayout}'."
            );

            return this.SetDocumentTabsLayout(
                targetLayout,
                $"Tabs Manager switches standard document tabs to {targetLayout}"
            )
                ? targetLayout
                : null;
        }

        private string? GetDocumentTabsLayout() {
            ThreadHelper.ThrowIfNotOnUIThread();

            try {
                var settingsWriter = GetSettingsWriter();
                if (settingsWriter == null) {
                    Helpers.Diagnostic.Logger.LogWarning("Unified settings writer is unavailable.");
                    return null;
                }

                var currentLayout = settingsWriter.GetValueOrThrow<string>(DocumentTabsLayoutSetting);

                Helpers.Diagnostic.Logger.LogDebug(
                    $"Read '{DocumentTabsLayoutSetting}' = '{currentLayout}'."
                );

                return currentLayout;
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError(
                    $"Failed to read '{DocumentTabsLayoutSetting}': {ex}"
                );

                return null;
            }
        }

        private bool SetDocumentTabsLayout(string layout, string commitReason) {
            ThreadHelper.ThrowIfNotOnUIThread();

            try {
                var settingsWriter = GetSettingsWriter();
                if (settingsWriter == null) {
                    Helpers.Diagnostic.Logger.LogWarning("Unified settings writer is unavailable.");
                    return false;
                }

                var currentLayout = settingsWriter.GetValueOrThrow<string>(DocumentTabsLayoutSetting);

                Helpers.Diagnostic.Logger.LogDebug(
                    $"Preparing '{DocumentTabsLayoutSetting}' change from '{currentLayout}' to '{layout}'."
                );

                if (string.Equals(currentLayout, layout, StringComparison.OrdinalIgnoreCase)) {
                    Helpers.Diagnostic.Logger.LogDebug("Requested tabs layout is already active.");
                    return true;
                }

                var changeResult = settingsWriter.EnqueueChange(DocumentTabsLayoutSetting, layout);
                Helpers.Diagnostic.Logger.LogDebug(
                    $"EnqueueChange result: outcome = '{changeResult.Outcome}', " +
                    $"will change effective value = {changeResult.CommitWillChangeEffectiveValue}, " +
                    $"message = '{changeResult.Message}'."
                );

                // EnqueueChange пока только добавляет изменение в очередь. Эти два результата означают,
                // что очередь сформирована успешно и теперь нужно вызвать RequestCommit для её применения.
                // На CommitWillChangeEffectiveValue не полагаемся: при переходе сверху налево Visual Studio
                // иногда возвращает false, хотя следующий RequestCommit действительно меняет расположение.
                if (changeResult.Outcome != SettingChangeOutcome.PendingCommit &&
                    changeResult.Outcome != SettingChangeOutcome.PendingCommitWithoutValidation) {
                    Helpers.Diagnostic.Logger.LogWarning(
                        $"Unable to enqueue '{DocumentTabsLayoutSetting}' change: " +
                        $"outcome = '{changeResult.Outcome}', message = '{changeResult.Message}'."
                    );

                    return false;
                }

                var commitResult = settingsWriter.RequestCommit(commitReason);
                Helpers.Diagnostic.Logger.LogDebug(
                    $"Changed '{DocumentTabsLayoutSetting}' from '{currentLayout}' to '{layout}': {commitResult}"
                );

                if (commitResult.Outcome != SettingCommitOutcome.Success) {
                    Helpers.Diagnostic.Logger.LogWarning(
                        $"Unable to commit '{DocumentTabsLayoutSetting}' change: " +
                        $"outcome = '{commitResult.Outcome}', message = '{commitResult.Message}'."
                    );

                    return false;
                }

                return true;
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError(
                    $"Failed to set '{DocumentTabsLayoutSetting}' to '{layout}': {ex}"
                );

                return false;
            }
        }

        /// <summary>
        /// Переключает отображение между оригинальным содержимым PART_TabListHost и кастомным контролом.
        /// </summary>
        /// <param name="enable">Если true — включить кастомные вкладки, иначе вернуть оригинал.</param>
        public void ToggleCustomTabs(bool enable) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope($"ToggleCustomTabs({enable})");

            ThreadHelper.ThrowIfNotOnUIThread();

            if (enable) {
                if (!this.IsCustomTabsEnabled) {
                    _documentTabsLayoutBeforeCustomTabs = this.GetDocumentTabsLayout();
                    this.IsCustomTabsEnabled = true;
                    this.EnsureStandardDocumentTabsOnLeft();

                    // Смена положения стандартных вкладок заставляет Visual Studio удалить старую область
                    // PART_TabListHost и создать новую. ApplicationIdle откладывает наш код до завершения
                    // этой перестройки, чтобы контрол Tabs Manager был вставлен уже в новую область.
                    Application.Current.Dispatcher.BeginInvoke(
                        new Action(
                            () => {
                                if (this.IsCustomTabsEnabled) {
                                    this.UpdateCustomTabsHost(true);
                                }
                            }
                        ),
                        DispatcherPriority.ApplicationIdle
                    );

                    return;
                }

                this.UpdateCustomTabsHost(true);
                return;
            }

            this.IsCustomTabsEnabled = false;
            // Сначала убираем Tabs Manager и возвращаем стандартные вкладки в текущую область.
            // После этого восстанавливаем их прежнее положение; Visual Studio при этом может удалить
            // текущую область вкладок и создать её заново.
            this.UpdateCustomTabsHost(false);

            var layoutToRestore = _documentTabsLayoutBeforeCustomTabs;
            _documentTabsLayoutBeforeCustomTabs = null;
            if (!string.IsNullOrWhiteSpace(layoutToRestore)) {
                this.SetDocumentTabsLayout(
                    layoutToRestore,
                    "Tabs Manager restores the previous document tabs layout"
                );
            }
        }

        private void UpdateCustomTabsHost(bool enable) {

            var mainWindow = Application.Current.MainWindow;
            if (mainWindow == null) {
                return;
            }

            var tabHost = Helpers.VisualTree.FindElementByName(mainWindow, "PART_TabListHost") as Decorator;
            if (tabHost == null) {
                Helpers.Diagnostic.Logger.LogWarning("PART_TabListHost not found");
                return;
            }

            // Если Decorator пересоздан — сбросим оригинальный контент
            if (_currentTabHost == null || !_currentTabHost.TryGetTarget(out var knownHost) || knownHost != tabHost) {
                _originalTabListHostContent = tabHost.Child;
                _currentTabHost = new WeakReference<Decorator>(tabHost);
            }

            if (enable) {
                if (tabHost.Child is Controls.TabsManagerToolWindowControl) {
                    return; // Уже вставлено
                }

                //Services.ExtensionServices.Initialize();

                var customControl = new Controls.TabsManagerToolWindowControl();
                customControl.Unloaded += this.OnInjectedControlUnloaded;

                tabHost.Child = customControl;
                _lastInjectedContent = new WeakReference<UIElement>(customControl);

                Helpers.Diagnostic.Logger.LogDebug("TabsManagerToolWindowControl injected.");
            }
            else {
                if (_originalTabListHostContent != null) {
                    tabHost.Child = _originalTabListHostContent;
                    _lastInjectedContent = null;

                    Helpers.Diagnostic.Logger.LogDebug("Restored original tab content.");

                    //Services.ExtensionServices.RequestShutdown();
                }
            }
        }

        /// <summary>
        /// Автоматическое переключение между оригинальным и кастомным таб-контролом.
        /// </summary>
        public void ToggleCustomTabs() {
            this.ToggleCustomTabs(!_isCustomTabsEnabled);
        }

        private void OnInjectedControlUnloaded(object sender, RoutedEventArgs e) {
            if (sender is FrameworkElement element) {
                element.Unloaded -= this.OnInjectedControlUnloaded;
            }

            Helpers.Diagnostic.Logger.LogDebug("TabsManagerToolWindowControl.Unloaded — re-evaluating state...");

            Application.Current.Dispatcher.BeginInvoke(new Action(() => {
                if (_isCustomTabsEnabled) {
                    this.ToggleCustomTabs(true); // повторно инжектим
                }
            }), DispatcherPriority.Background);
        }
    }
}

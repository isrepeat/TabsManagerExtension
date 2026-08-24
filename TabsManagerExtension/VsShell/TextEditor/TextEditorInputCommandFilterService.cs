using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Collections.Generic;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using TabsManagerExtension.VsShell.Document.Services;


namespace TabsManagerExtension.VsShell.TextEditor.Services {
    /// <summary>
    /// Серивис, автоматически отслеживающий активный редактор и переустанавливающий фильтр команд.
    /// Внешние подписчики могут подписаться на события команд один раз — фильтр будет переустановлен при смене редактора.
    /// </summary>
    public sealed class TextEditorInputCommandFilterService :
        TabsManagerExtension.Services.SingletonServiceBase<TextEditorInputCommandFilterService>,
        TabsManagerExtension.Services.IExtensionService {

        private static readonly VSConstants.VSStd2KCmdID[] _trackedStd2Commands = new[] {
            VSConstants.VSStd2KCmdID.TAB,
            VSConstants.VSStd2KCmdID.UP,
            // VS отправляет Shift+Arrow отдельными extend-командами, а не обычными UP/DOWN.
            VSConstants.VSStd2KCmdID.UP_EXT,
            VSConstants.VSStd2KCmdID.DOWN,
            VSConstants.VSStd2KCmdID.DOWN_EXT,
            VSConstants.VSStd2KCmdID.LEFT,
            VSConstants.VSStd2KCmdID.RIGHT,
            VSConstants.VSStd2KCmdID.RETURN,
            VSConstants.VSStd2KCmdID.CANCEL,
            VSConstants.VSStd2KCmdID.DELETE,
            VSConstants.VSStd2KCmdID.BACKSPACE,
            // Space приходит как TYPECHAR, поэтому фильтру нужен исходный OLE-аргумент символа.
            VSConstants.VSStd2KCmdID.TYPECHAR,
        };

        private static readonly VSConstants.VSStd97CmdID[] _trackedStd97Commands = new[] {
            VSConstants.VSStd97CmdID.Delete,
            VSConstants.VSStd97CmdID.Cancel,
            VSConstants.VSStd97CmdID.Escape,
            VSConstants.VSStd97CmdID.SelectAll,
            VSConstants.VSStd97CmdID.Undo,
        };

        private IVsTextView? _currentTextView;
        private TextEditorCommandFilter? _currentFilter;
        private bool _installScheduled;

        private readonly HashSet<FrameworkElement> _trackedElements = new();
        private FrameworkElement? _lastInputTarget;
        // В edit mode команды должны идти во вкладки даже после активации editor frame.
        private FrameworkElement? _forcedInputTarget;
        private Func<Key, bool>? _forcedInputKeyAvailability;

        public TextEditorInputCommandFilterService() { }

        //
        // IExtensionService
        //
        public IReadOnlyList<Type> DependsOn() {
            return new[] {
                typeof(VsShell.Document.Services.VsDocumentActivationTrackerService),
                typeof(VsShell.Solution.Services.VsWindowFrameActivationTrackerService),
            };
        }


        public void Initialize() {
            ThreadHelper.ThrowIfNotOnUIThread();

            VsShell.Document.Services.VsDocumentActivationTrackerService.Instance.OnDocumentActivated += this.OnDocumentActivatedExternally;
            VsShell.Solution.Services.VsWindowFrameActivationTrackerService.Instance.VsWindowFrameActivated += this.OnVsWindowFrameActivated;
            
            this.InstallToActiveEditor();

            Helpers.Diagnostic.Logger.LogDebug("[TextEditorInputCommandFilterService] Initialized.");
        }


        public void Shutdown() {
            ThreadHelper.ThrowIfNotOnUIThread();

            VsShell.Solution.Services.VsWindowFrameActivationTrackerService.Instance.VsWindowFrameActivated -= this.OnVsWindowFrameActivated;
            VsShell.Document.Services.VsDocumentActivationTrackerService.Instance.OnDocumentActivated -= this.OnDocumentActivatedExternally;
            this.UninstallFilter();
            ClearInstance();

            Helpers.Diagnostic.Logger.LogDebug("[TextEditorInputCommandFilterService] Shutdown.");
        }


        //
        // Api
        //
        public void AddTrackedInputElement(FrameworkElement element) {
            if (element == null || _trackedElements.Contains(element)) {
                return;
            }

            _trackedElements.Add(element);

            element.GotKeyboardFocus += OnTargetGotFocus;
            element.LostKeyboardFocus += OnTargetLostFocus;

            if (element.IsKeyboardFocusWithin) {
                this.Enable();
                _lastInputTarget = element;
            }
        }


        public void RemoveTrackedInputElement(FrameworkElement element) {
            if (element == null || !_trackedElements.Contains(element)) {
                return;
            }

            _trackedElements.Remove(element);

            element.GotKeyboardFocus -= OnTargetGotFocus;
            element.LostKeyboardFocus -= OnTargetLostFocus;

            if (ReferenceEquals(_lastInputTarget, element)) {
                _lastInputTarget = null;
            }

            if (ReferenceEquals(_forcedInputTarget, element)) {
                _forcedInputTarget = null;
            }

            if (_forcedInputTarget == null && _lastInputTarget == null) {
                this.Disable();
            }
        }


        public void SetForcedInputTarget(FrameworkElement? element, Func<Key, bool>? keyAvailability = null) {
            if (element != null && !_trackedElements.Contains(element)) {
                throw new InvalidOperationException("Forced input target must be registered first.");
            }

            // Forced target имеет приоритет над фактическим WPF-фокусом. Это позволяет оставить
            // активным документ Visual Studio, но обрабатывать навигационные клавиши панелью вкладок.
            _forcedInputTarget = element;
            _forcedInputKeyAvailability = element == null ? null : keyAvailability;
            if (_currentFilter != null) {
                _currentFilter.IsSpaceCommandEnabled = _forcedInputTarget != null;
                _currentFilter.AreNavigationCommandsEnabled = _forcedInputTarget != null;
            }

            if (_forcedInputTarget != null || _lastInputTarget?.IsKeyboardFocusWithin == true) {
                this.Enable();
            }
            else {
                this.Disable();
            }
        }


        public void Enable() {
            if (_currentFilter != null) {
                _currentFilter.IsEnabled = true;
            }
        }


        public void Disable() {
            if (_currentFilter != null) {
                _currentFilter.IsEnabled = false;
            }
        }


        //
        // Event handlers
        //
        private void OnDocumentActivatedExternally(_EventArgs.DocumentNavigationEventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.ScheduleInstallToActiveEditor();
        }


        private void OnVsWindowFrameActivated(IVsWindowFrame vsWindowFrame) {
            ThreadHelper.ThrowIfNotOnUIThread();

            this.ScheduleInstallToActiveEditor();
        }

        private void ScheduleInstallToActiveEditor() {
            if (_installScheduled) {
                return;
            }

            _installScheduled = true;
            VsixThreadHelper.RunOnUiThread(Dispatcher.CurrentDispatcher, () => {
                _installScheduled = false;
                this.InstallToActiveEditor();
            }, DispatcherPriority.Background);
        }


        private void OnTargetGotFocus(object sender, RoutedEventArgs e) {
            this.Enable();
            _lastInputTarget = sender as FrameworkElement;
        }


        private void OnTargetLostFocus(object sender, RoutedEventArgs e) {
            _lastInputTarget = null;

            // Потеря обычного WPF-фокуса не выключает фильтр, пока действует edit mode.
            if (_forcedInputTarget == null) {
                this.Disable();
            }
        }


        private void OnCommandIntercepted(Guid cmdGroup, uint cmdId, IntPtr inputArgument) {
            // В edit mode используем принудительную цель; вне его перенаправляем ввод только
            // действительно сфокусированному зарегистрированному контролу.
            var inputTarget = _forcedInputTarget ?? _lastInputTarget;
            if (inputTarget == null || (_forcedInputTarget == null && !inputTarget.IsKeyboardFocusWithin)) {
                return;
            }

            if (_currentFilter != null) {
                var key = _currentFilter.TryResolveKey(cmdGroup, cmdId, inputArgument);
                if (key.HasValue) {
                    if (key == Key.A || key == Key.Z) {
                        Helpers.Diagnostic.Logger.LogDebug($"[NavigationInput] OLE command resolved as {key.Value}; forced={_forcedInputTarget != null}, panelFocused={inputTarget.IsKeyboardFocusWithin}.");
                    }

                    // Forced target хранит панель для navigation mode, но её дочерний TextBox
                    // может иметь настоящий WPF-фокус во время rename. В таком случае посылаем
                    // событие прямо полю, иначе Enter/стрелки попадут в корень панели.
                    var routedInputTarget = _forcedInputTarget != null &&
                                            inputTarget.IsKeyboardFocusWithin &&
                                            Keyboard.FocusedElement is FrameworkElement focusedElement
                        ? focusedElement
                        : inputTarget;

                    this.RedirectKeyInput(routedInputTarget, key.Value);
                }
            }
        }

        private void OnCommandPassedThrough(Guid cmdGroup, uint cmdId, IntPtr inputArgument) {
            // ...
        }


        //
        // Internal logic
        //
        private void InstallToActiveEditor() {
            //using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("TextEditorCommandFilterService.InstallToActiveEditor()");

            if (PackageServices.VsTextManager.GetActiveView(1, null, out var newView) == VSConstants.S_OK) {
                if (newView != null) {
                    if (!ReferenceEquals(_currentTextView, newView)) {
                        this.UninstallFilter();
                        this.InstallFilterToView(newView);
                    }
                }
            }
        }


        private void InstallFilterToView(IVsTextView view) {
            //using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("TextEditorCommandFilterService.InstallFilterToView()");
            ThreadHelper.ThrowIfNotOnUIThread();

            var filter = new TextEditorCommandFilter(_trackedStd2Commands, _trackedStd97Commands);
            // TYPECHAR перехватывается только в edit mode, чтобы обычный ввод текста не затрагивался.
            filter.IsSpaceCommandEnabled = _forcedInputTarget != null;
            filter.AreNavigationCommandsEnabled = _forcedInputTarget != null;
            filter.CanInterceptNavigationKey = this.CanInterceptNavigationKey;
            int result = view.AddCommandFilter(filter, out IOleCommandTarget next);

            if (result == VSConstants.S_OK) {
                filter.SetNext(next);
                filter.CommandPassedThrough += this.OnCommandPassedThrough;
                filter.CommandIntercepted += this.OnCommandIntercepted;

                _currentFilter = filter;
                _currentTextView = view;

                // Новый editor view получает новый command filter; восстанавливаем его состояние
                // после смены документа, не ожидая дополнительного события WPF-фокуса.
                if (_forcedInputTarget != null || _lastInputTarget?.IsKeyboardFocusWithin == true) {
                    _currentFilter.IsEnabled = true;
                }
            }
            else {
                Helpers.Diagnostic.Logger.LogWarning($"[TextEditorCommandFilterController] Failed to install the filter. HRESULT = 0x{result:X8}");
            }
        }


        private void UninstallFilter() {
            //using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("TextEditorCommandFilterService.UninstallFilter()");

            if (_currentFilter != null) {
                _currentFilter.IsEnabled = false;
                _currentFilter.CommandIntercepted -= this.OnCommandIntercepted;
                _currentFilter.CommandPassedThrough -= this.OnCommandPassedThrough;
            }
            if (_currentTextView != null) {
                _currentTextView.RemoveCommandFilter(_currentFilter);
            }
            _currentTextView = null;
            _currentFilter = null;
        }

        private bool CanInterceptNavigationKey(Key key) {
            return _forcedInputTarget != null && (_forcedInputKeyAvailability?.Invoke(key) ?? true);
        }



        private void RedirectKeyInput(FrameworkElement target, Key key) {
            var inputEvent = new KeyEventArgs(
                Keyboard.PrimaryDevice,
                PresentationSource.FromVisual(target),
                Environment.TickCount,
                key);
            inputEvent.RoutedEvent = Keyboard.KeyDownEvent;

            // RaiseEvent фиксирует routed target явно. Через InputManager синтетическая команда
            // могла повторно выбрать editor view, несмотря на WPF-фокус inline TextBox.
            target.RaiseEvent(inputEvent);
        }
    }
}

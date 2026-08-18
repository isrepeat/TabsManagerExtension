using System;
using System.Linq;
using System.Windows.Threading;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.VisualStudio.Shell;


namespace TabsManagerExtension.Services {
    public sealed class ExtensionStatusService :
        SingletonServiceBase<ExtensionStatusService>,
        IExtensionService {

        public const string IncludeGraphFeature = "IncludeGraph";

        public ObservableCollection<State.BackgroundOperationStatus> Operations { get; } = new();
        public event Action<string, bool>? FeatureReadinessChanged;

        private readonly Dictionary<string, bool> _featureReadiness = new(StringComparer.Ordinal);
        private Dispatcher _dispatcher = null!;

        public IReadOnlyList<Type> DependsOn() {
            return Array.Empty<Type>();
        }

        public void Initialize() {
            ThreadHelper.ThrowIfNotOnUIThread();
            _dispatcher = Dispatcher.CurrentDispatcher;
            _featureReadiness[IncludeGraphFeature] = false;
            Helpers.Diagnostic.Logger.LogDebug("[ExtensionStatusService] Initialized.");
        }

        public void Shutdown() {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.Operations.Clear();
            _featureReadiness.Clear();
            this.FeatureReadinessChanged = null;
            ClearInstance();
            Helpers.Diagnostic.Logger.LogDebug("[ExtensionStatusService] Disposed.");
        }

        public bool IsFeatureReady(string feature) {
            return _featureReadiness.TryGetValue(feature, out bool isReady) && isReady;
        }

        public void SetFeatureReady(string feature, bool isReady) {
            this.RunOnUiThread(() => {
                if (this.IsFeatureReady(feature) == isReady) {
                    return;
                }

                _featureReadiness[feature] = isReady;
                this.FeatureReadinessChanged?.Invoke(feature, isReady);
            });
        }

        public void ReportOperation(string key, string title, string stage, double progress) {
            this.RunOnUiThread(() => {
                var operation = this.Operations.FirstOrDefault(item => item.Key == key);
                if (operation == null) {
                    double normalizedProgress = Math.Max(0, Math.Min(100, progress));
                    operation = new State.BackgroundOperationStatus(key, title, stage, normalizedProgress);
                    this.Operations.Add(operation);
                    return;
                }

                operation.Title = title;
                operation.Stage = stage;
                operation.Progress = Math.Max(0, Math.Min(100, progress));
            });
        }

        public void RemoveOperation(string key) {
            this.RunOnUiThread(() => {
                var operation = this.Operations.FirstOrDefault(item => item.Key == key);
                if (operation != null) {
                    this.Operations.Remove(operation);
                }
            });
        }

        private void RunOnUiThread(Action action) {
            if (_dispatcher.CheckAccess()) {
                action();
            }
            else {
                VsixThreadHelper.RunOnUiThread(_dispatcher, action, DispatcherPriority.Background);
            }
        }
    }
}

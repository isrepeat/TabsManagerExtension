using Helpers;


namespace TabsManagerExtension.State {
    public sealed class BackgroundOperationStatus : ObservableObject {
        public string Key { get; }

        private string _title;
        public string Title {
            get => _title;
            internal set => base.SetProperty(ref _title, value);
        }

        private string _stage;
        public string Stage {
            get => _stage;
            internal set => base.SetProperty(ref _stage, value);
        }

        private double _progress;
        public double Progress {
            get => _progress;
            internal set => base.SetProperty(ref _progress, value);
        }

        public BackgroundOperationStatus(string key, string title, string stage, double progress) {
            this.Key = key;
            _title = title;
            _stage = stage;
            _progress = progress;
        }
    }
}

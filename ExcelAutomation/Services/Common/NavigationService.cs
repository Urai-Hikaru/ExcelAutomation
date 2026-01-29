using CommunityToolkit.Mvvm.ComponentModel;
using ExcelAutomation.ViewModels;

namespace ExcelAutomation.Services.Common
{
    public class NavigationService : ObservableObject
    {
        // シングルトン
        public static NavigationService Instance { get; } = new NavigationService();

        private NavigationService() { }

        // 現在表示中のViewModel
        private BaseViewModel? _currentView;
        public BaseViewModel? CurrentView
        {
            get => _currentView;
            set => SetProperty(ref _currentView, value);
        }

        // ViewModelのキャッシュ（一度作った画面はここに保持する）
        private readonly Dictionary<string, BaseViewModel> _viewModelCache = new();

        public void NavigateTo(string key)
        {
            if (!_viewModelCache.ContainsKey(key))
            {
                switch (key)
                {
                    case "DataProcessingView":
                        _viewModelCache[key] = new DataProcessingViewModel();
                        break;
                    case "BatchProcessingView":
                        _viewModelCache[key] = new BatchProcessingViewModel();
                        break;
                    default:
                        throw new ArgumentException("Unknown view key", nameof(key));
                }
            }

            // キャッシュから取り出してセット
            CurrentView = _viewModelCache[key];
        }
    }
}
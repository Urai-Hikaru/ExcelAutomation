using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ExcelAutomation.Services;
using System.Windows.Input;

namespace ExcelAutomation.ViewModels
{
    public class MainViewModel : ObservableObject
    {
        // ナビゲーションサービスへの参照
        public NavigationService Navigation { get; } = NavigationService.Instance;

        // ステータスサービスへの参照
        public StatusService Status { get; } = StatusService.Instance;

        // ログサービスへの参照
        public SystemLogService SystemLog { get; } = SystemLogService.Instance;

        public ViewModelBase? CurrentView => Navigation.CurrentView;

        public ICommand NavigateCommand { get; }

        public MainViewModel()
        {
            // NavigationServiceのプロパティ変更を監視して、MainViewModel側にも通知する
            Navigation.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(NavigationService.CurrentView))
                {
                    OnPropertyChanged(nameof(CurrentView));
                }
            };

            // コマンド実行時はServiceに依頼するだけ
            NavigateCommand = new RelayCommand<string>(key =>
            {
                if (key != null) Navigation.NavigateTo(key);
            });

            // 初期画面
            Navigation.NavigateTo("DataProcessingView");
        }
    }
}
using ExcelAutomation.ViewModels;

namespace ExcelAutomation.Services
{
    // アプリ全体で共有する状態管理クラス
    public class StatusService : ViewModelBase
    {
        // シングルトン
        public static StatusService Instance { get; } = new StatusService();

        private StatusService() { }

        // ==========================================
        // 共有プロパティ
        // ==========================================

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set => SetProperty(ref _isBusy, value);
        }

        private string _statusMessage = "待機中";
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        private int _progressValue;
        public int ProgressValue
        {
            get => _progressValue;
            set => SetProperty(ref _progressValue, value);
        }

        // ==========================================
        // 便利メソッド
        // ==========================================
        public void Start(string message = "処理中...")
        {
            IsBusy = true;
            StatusMessage = message;
            ProgressValue = 0;
        }

        public void End(string message = "完了")
        {
            IsBusy = false;
            StatusMessage = message;
            ProgressValue = 100;
        }

        /// <summary>
        /// 進捗状況を更新
        /// </summary>
        /// <param name="value">進捗率(0-100)</param>
        /// <param name="message">メッセージ(省略可、nullの場合は変更なし)</param>
        public void UpdateProgress(int value, string? message = null)
        {
            ProgressValue = value;

            // メッセージが指定されている場合のみ更新
            if (message != null)
            {
                StatusMessage = message;
            }
        }
    }
}
using System.Windows.Input;

namespace ExcelAutomation.Commands
{
    /// <summary>
    /// コマンド実行ロジックをデリゲートで受け取る汎用コマンドクラス
    /// </summary>
    public class RelayCommand : ICommand
    {
        private readonly Func<Task> _execute;
        private readonly Func<bool>? _canExecute;

        // コンストラクタ
        public RelayCommand(Func<Task> execute, Func<bool>? canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        // ICommandの実装
        public event EventHandler? CanExecuteChanged;

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);

        public bool CanExecute(object? parameter) => _canExecute == null || _canExecute();

        // 非同期で実行
        public async void Execute(object? parameter) => await _execute();
    }
}
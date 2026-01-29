using ExcelAutomation.Commands;
using ExcelAutomation.Data;
using ExcelAutomation.Services.Api;

namespace ExcelAutomation.ViewModels
{
    public class LoginViewModel : BaseViewModel
    {
        private readonly IAuthService _authService;

        public LoginViewModel(IAuthService authService)
        {
            _authService = authService;
            LoginCommand = new RelayCommand(ExecuteLoginAsync, CanExecuteLogin);
        }

        private string _username = string.Empty;
        public string Username
        {
            get => _username;
            set
            {
                if (SetProperty(ref _username, value))
                    LoginCommand.RaiseCanExecuteChanged();
            }
        }

        private string _password = string.Empty;
        public string Password
        {
            get => _password;
            set
            {
                if (SetProperty(ref _password, value))
                    LoginCommand.RaiseCanExecuteChanged();
            }
        }

        private string _errorMessage = string.Empty;
        public string ErrorMessage
        {
            get => _errorMessage;
            set => SetProperty(ref _errorMessage, value);
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                if (SetProperty(ref _isBusy, value))
                    LoginCommand.RaiseCanExecuteChanged();
            }
        }

        public event Action<bool>? RequestClose;
        public RelayCommand LoginCommand { get; }

        private bool CanExecuteLogin()
        {
            return !IsBusy
                   && !string.IsNullOrWhiteSpace(Username)
                   && !string.IsNullOrWhiteSpace(Password);
        }

        // ログイン実行
        private async Task ExecuteLoginAsync()
        {
            try
            {
                IsBusy = true;
                ErrorMessage = string.Empty;

                var result = await _authService.LoginAsync(Username, Password);

                if (result.IsSuccess && result.Data != null)
                {
                    // データがある場合はセッションにセット
                    UserSession.Set(result.Data);

                    // 成功通知
                    RequestClose?.Invoke(true);
                }
                else
                {
                    ErrorMessage = result.ErrorMessage ?? "ログインに失敗しました。";
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"システムエラーが発生しました: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
                LoginCommand.RaiseCanExecuteChanged();
            }
        }
    }
}
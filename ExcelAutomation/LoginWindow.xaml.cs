using ExcelAutomation.Services.Api;
using ExcelAutomation.ViewModels;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelAutomation.Views
{
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();

            var authService = new AuthService();
            var viewModel = new LoginViewModel(authService);

            viewModel.RequestClose += (isSuccess) =>
            {
                // ダイアログの結果を設定
                this.DialogResult = isSuccess;
                this.Close();
            };

            this.DataContext = viewModel;

            UsernameBox.Focus();
        }

        // パスワード連携ロジック
        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (this.DataContext is LoginViewModel viewModel)
            {
                viewModel.Password = ((PasswordBox)sender).Password;
            }
        }

        // ウィンドウのドラッグ移動
        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        // 終了ボタン
        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
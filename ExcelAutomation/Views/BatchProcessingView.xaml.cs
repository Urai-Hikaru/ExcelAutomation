using ExcelAutomation.Models;
using ExcelAutomation.ViewModels;
using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace ExcelAutomation.Views
{
    public partial class BatchProcessingView : UserControl
    {
        public BatchProcessingView()
        {
            InitializeComponent();
        }

        // フォルダ追加ボタン押下時の処理
        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog
            {
                Title = "処理対象のフォルダを選択してください",
                Multiselect = false,
                ValidateNames = true
            };

            // ダイアログを表示
            if (dialog.ShowDialog() == true)
            {
                string selectedPath = dialog.FolderName;
                LoadExcelFiles(selectedPath);
            }
        }

        // エクセルファイルを抽出し、全選択状態で追加
        private void LoadExcelFiles(string folderPath)
        {
            try
            {
                if (!(this.DataContext is BatchProcessingViewModel vm))
                {
                    MessageBox.Show("ViewModelが正しく設定されていません。");
                    return;
                }

                vm.FileList.Clear();

                // 対象の拡張子
                string[] extensions = { ".xls", ".xlsx" };

                // フォルダ内のファイルをスキャン
                var files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                                     .Where(f => extensions.Contains(Path.GetExtension(f).ToLower()));

                int newCount = 0;

                foreach (var file in files)
                {
                    vm.FileList.Add(new FileItemModel
                    {
                        IsSelected = true,
                        FileName = Path.GetFileName(file),
                        FilePath = file,
                    });
                    newCount++;
                }

                if (newCount == 0)
                {
                    MessageBox.Show("指定されたフォルダにExcelファイルが見つかりませんでした。", "確認", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ファイルの読み込み中にエラーが発生しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
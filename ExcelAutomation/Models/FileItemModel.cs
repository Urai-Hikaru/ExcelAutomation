using CommunityToolkit.Mvvm.ComponentModel;

namespace ExcelAutomation.Models
{
    public class FileItemModel : ObservableObject
    {
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        public string FileName { get; set; } = string.Empty;

        public string FilePath { get; set; } = string.Empty;
    }
}
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace ExcelAutomation.Views
{
    public partial class DataProcessingView : UserControl
    {
        public DataProcessingView()
        {
            InitializeComponent();
        }

        private void DatePicker_CalendarOpened(object sender, RoutedEventArgs e)
        {
            if (sender is DatePicker datePicker)
            {
                // カレンダーコントロールを取得してモードを変更
                var popup = datePicker.Template.FindName("PART_Popup", datePicker) as Popup;
                if (popup != null && popup.Child is Calendar calendar)
                {
                    calendar.DisplayMode = CalendarMode.Year;
                }
            }
        }
    }
}
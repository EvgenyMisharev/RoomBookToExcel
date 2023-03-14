using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace RoomBookToExcel
{
    public partial class RoomBookToExcelWPF : Window
    {
        public string ExportOptionName;
        public RoomBookToExcelWPF()
        {
            InitializeComponent();
        }
        private void btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            ExportOptionName = (groupBox_ExportOption.Content as System.Windows.Controls.Grid)
                .Children.OfType<RadioButton>()
                .FirstOrDefault(rb => rb.IsChecked.Value == true)
                .Name;
            DialogResult = true;
            Close();
        }

        private void btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        private void RoomBookToExcelWPF_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Space)
            {
                ExportOptionName = (groupBox_ExportOption.Content as System.Windows.Controls.Grid)
                    .Children.OfType<RadioButton>()
                    .FirstOrDefault(rb => rb.IsChecked.Value == true)
                    .Name;
                DialogResult = true;
                Close();
            }

            else if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }
    }
}

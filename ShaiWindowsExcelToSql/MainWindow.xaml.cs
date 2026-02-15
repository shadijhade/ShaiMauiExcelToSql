using System.Windows;
using ShaiWindowsExcelToSql.ViewModels;

namespace ShaiWindowsExcelToSql
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
        }
    }
}
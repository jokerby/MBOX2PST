using System.Windows;
using System.Windows.Input;

namespace Converter
{
    public partial class MainWindow : Window
    {
        private readonly ViewModel _vm = new ViewModel();

        public MainWindow()
        {
            InitializeComponent();
            DataContext = _vm;
        }

        private void TextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            _vm.SelectSourcePath();
        }

        private void TextBox_PreviewMouseDown_1(object sender, MouseButtonEventArgs e)
        {
            _vm.SelectDestinationPath();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _vm.Run();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            _vm.Stop();
        }
    }
}
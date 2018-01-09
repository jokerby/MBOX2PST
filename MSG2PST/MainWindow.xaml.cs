using System;
using System.Collections.Generic;
using System.Windows;

namespace MSG2PST
{
    public partial class MainWindow : Window
    {
        private readonly ViewModel _vm;

        public MainWindow()
        {
            try
            {
                _vm = new ViewModel();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            InitializeComponent();
            DataContext = _vm;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _vm.Run();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            _vm.Stop();
        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }

        private void TextBox_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _vm.SelectSourcePath();
        }

        private void TextBox_PreviewMouseDown_1(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _vm.SelectDestinationPath();
        }
    }
}

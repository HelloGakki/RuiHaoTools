using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using RuiHaoConvertor.ViewModel;

namespace RuiHaoConvertor
{
    /// <summary>
    /// ConvertorClient.xaml 的交互逻辑
    /// </summary>
    public partial class ConvertorClient : Window
    {
        public ConvertorClient()
        {
            InitializeComponent();
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            //var model = DataContext as BOMConvertor;
            //model.Dispose();
            this.Close();
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.BomConvertor.DelayFileConvertor();
        }

        private void addButton_Click(object sender, RoutedEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.BomConvertor.GetFilePath();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.BomConvertor.Dispose();
        }

        private void BOMMessage_MouseLeave(object sender, MouseEventArgs e)
        {
            var control = sender as ScrollViewer;
            control.ScrollToEnd();
        }

        private void resistorRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.RCConvertor.Components = Components.Resistance;
        }

        private void capacitorRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.RCConvertor.Components = Components.Capacitance;
        }

        private void confirmButton_Click(object sender, RoutedEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.RCConvertor.Coding();
        }
    }
}

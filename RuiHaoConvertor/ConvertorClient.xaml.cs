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
        /// <summary>
        /// 退出窗口引发关闭事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            //var model = DataContext as BOMConvertor;
            //model.Dispose();
            this.Close();
        }
        /// <summary>
        /// 保存文件按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            if(tabControl.SelectedIndex == 0)
                model.BomConvertor.DelayFileConvertor();
            if (tabControl.SelectedIndex == 1)
                model.RCConvertor.DelaySaveCode();
            if (tabControl.SelectedIndex == 2)
                model.CodeConvertor.Save();
        }
        /// <summary>
        /// BOM导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void addButton_Click(object sender, RoutedEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.BomConvertor.GetFilePath();
        }
        /// <summary>
        /// 关闭窗口引发的事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.BomConvertor.Dispose();
            model.CodeConvertor.Dispose();
            model.RCConvertor.Dispose();
        }
        /// <summary>
        /// 保持滚动条在最底端
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BOMMessage_MouseLeave(object sender, MouseEventArgs e)
        {
            var control = sender as ScrollViewer;
            control.ScrollToEnd();
        }
        /// <summary>
        /// 电阻电容选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void resistorRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.RCConvertor.Components = Components.Resistance;
        }
        /// <summary>
        /// 电阻电容选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void capacitorRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.RCConvertor.Components = Components.Capacitance;
        }
        /// <summary>
        /// 转换件号按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void confirmButton_Click(object sender, RoutedEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.RCConvertor.Coding();
        }
        /// <summary>
        /// 转换件号滚动条保持最底端
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rcScrollViewer_MouseLeave(object sender, MouseEventArgs e)
        {
            var control = sender as ScrollViewer;
            control.ScrollToEnd();
        }
        /// <summary>
        /// 直接拖入文件的方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabItem_DragEnter(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return;
             string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            pathTextBox.Text = files[0];
        }
        /// <summary>
        /// 保持滚动条在最底端
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void codeScrollViewer_MouseLeave(object sender, MouseEventArgs e)
        {
            var control = sender as ScrollViewer;
            control.ScrollToEnd();
        }
        /// <summary>
        /// 转换件号开始
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void convertButton_Click(object sender, RoutedEventArgs e)
        {
            var model = DataContext as ConvertorViewModel;
            model.CodeConvertor.DelayConvertorStart();
        }
        /// <summary>
        /// 命令可执行性
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CommandCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            if (e.Command == ConvertorViewModel.ShowAboutCommand)
                e.CanExecute = true;
        }
        /// <summary>
        /// 命令执行操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CommandExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            if (e.Command == ConvertorViewModel.ShowAboutCommand)
            {
                var aboutWin = new About();
                aboutWin.ShowDialog();
            }
        }
    }
}

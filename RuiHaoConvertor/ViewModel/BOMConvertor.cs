using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using RuiHaoConvertor;
using System.Windows.Forms;
using System.Threading;

namespace RuiHaoConvertor.ViewModel
{
    public class BOMConvertor : INotifyPropertyChanged
    {
        #region " 变更通知接口实现 "

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChange(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region " 私有变量定义 "

        private string _message, _filePath;
        ExcelHelper sourceExcel, exportExcel;
        bool _isConvertor;

        #endregion

        #region " 绑定属性定义 "

        public string Message
        {
            get
            {
                return _message;
            }

            set
            {
                _message = value;
                OnPropertyChange("Message");
            }
        }

        public string FilePath
        {
            get
            {
                return _filePath;
            }

            set
            {
                _filePath = value;
                OnPropertyChange("FilePath");
            }
        }

        public bool IsConvertor
        {
            get
            {
                return _isConvertor;
            }

            set
            {
                _isConvertor = value;
                OnPropertyChange("IsConvertor");
            }
        }

        #endregion

        #region " 构造函数 "

        public BOMConvertor()
        {
            //sourceExcel = new ExcelHelper();
            //exportExcel = new ExcelHelper();
            Message = "感谢使用本软件\r\n";
            _filePath = "";
            IsConvertor = false;
        }

        #endregion

        #region " 方法定义 "
        /// <summary>
        /// 获取文件路径
        /// </summary>
        public void GetFilePath()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择文件";
            openFileDialog.Filter = "Excel Worksheets|*.xlsx|Excel Worksheets|*.xls";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
                FilePath = openFileDialog.FileName;
            IsConvertor = true;
        }
        private void FileConvertor()
        {
            try
            {
                // 创建excel客户端
                if (sourceExcel == null)
                    sourceExcel = new ExcelHelper();
                if (exportExcel == null)
                    exportExcel = new ExcelHelper();

                // 打开文件
                sourceExcel.Show();
                exportExcel.Show();
                sourceExcel.Open(FilePath);
                exportExcel.Open(Environment.CurrentDirectory + @"/" + "模板.xlsx");
                sourceExcel.Hide();
                exportExcel.Hide();
                sourceExcel.SetActivitySheet("Export");
                exportExcel.SetActivitySheet("Export");

                // 输出状态
                Message += "打开文件：" + FilePath + "成功" + "\r\n";
                Message += "正在导出...\r\n";

                // 读数据，填数据
                int count = 0;
                for (var x = 1; x <= 10; x++)
                {
                    List<string> dataList = new List<string>();

                    for (var y = 1; ; y++)
                    {
                        object data = sourceExcel.GetCellValue(y, x);

                        if (data == null)
                            break;
                        else
                            dataList.Add(data.ToString());
                    }
                    if (dataList.Count == 0)
                        continue;
                    int columnIndex = GetCoordinate(dataList[0]);
                    if (columnIndex == -1)
                        continue;
                    dataList.RemoveAt(0);
                    if (dataList.Count > count)
                        count = dataList.Count;
                    for (int rowIndex = 0; rowIndex < dataList.Count; rowIndex++)
                    {
                        if (columnIndex == 11)
                            exportExcel.SetCellValue(rowIndex + 7, columnIndex, Convert.ToDouble(dataList[rowIndex]));
                        else
                        {
                            exportExcel.SetCellValue(rowIndex + 7, columnIndex, dataList[rowIndex]);
                            // 合并单元格
                            if (columnIndex == 3)
                                exportExcel.SetMergeCells(exportExcel.GetRange(rowIndex + 7, columnIndex, rowIndex + 7, columnIndex + 1));
                            if (columnIndex == 6)
                                exportExcel.SetMergeCells(exportExcel.GetRange(rowIndex + 7, columnIndex, rowIndex + 7, columnIndex + 2));
                            if (columnIndex == 12)
                                exportExcel.SetMergeCells(exportExcel.GetRange(rowIndex + 7, columnIndex, rowIndex + 7, columnIndex + 3));
                        }
                    }
                }

                // 填充序列号
                for (var index = 1; index <= count; index++)
                {
                    exportExcel.SetCellValue(7 + index - 1, 1, index);
                }

                // 设置样式
                exportExcel.Copy(exportExcel.GetRange("R1", "AG10"), exportExcel.GetRange(count + 7, 1, 10 + count, 16));   // 复制变更信息到对应位置
                exportExcel.DeleteColumn("R", "AG");    // 删除变更信息
                exportExcel.SetCellValue("D2", "应凌峰");
                exportExcel.SetCellBorder(7, 1, 7 + count - 1, 16); // 添加边框
                exportExcel.setCellTextByFormat(exportExcel.GetRange(7, 1, 7 + count - 1 + 10, 16), "微软雅黑", "10");
                exportExcel.SetCellValue(count - 1 + 3 + 7, 10, "应凌峰");
                exportExcel.SetCellValue(count - 1 + 3 + 7, 11, DateTime.Now.ToShortDateString());
                exportExcel.setCellTextByFormat(exportExcel.GetRange(count - 1 + 3 + 7, 11), "微软雅黑", "10", stringFormat: "yyyy-m-d");


                // 保存文件
                exportExcel.Save("Export");
                exportExcel.Saved();
                sourceExcel.Saved();
                exportExcel.Close();
                sourceExcel.Close();

                // 输出状态
                Message += "导出BOM成功" + "\r\n" + "共" + count.ToString() + "项" + "\r\n";
            }
            catch (Exception e)
            {
                Message += "转换失败:" + e.Message + "\r\n";
                exportExcel.Saved();
                sourceExcel.Saved();
                exportExcel.Close();
                sourceExcel.Close();
            }
        }
        private int GetCoordinate(string title)
        {
            switch (title)
            {
                case "Text Field1":
                    return 2;
                case "Text Field2":
                    return 3;
                case "Text Field4":
                    return 5;
                case "Text Field3":
                    return 6;
                case "Text Field5":
                    return 10;
                case "Quantity":
                    return 11;
                case "Text Field6":
                    return 9;
                case "Designator":
                    return 12;
                default:
                    return -1;
            }
        }

        public void Dispose()
        {
            if (exportExcel != null)
                exportExcel.Dispose();
            if (sourceExcel != null)
                sourceExcel.Dispose();
        }

        public void DelayFileConvertor()
        {
            Thread fileConvertorThread = new Thread(new ThreadStart(FileConvertor));
            fileConvertorThread.IsBackground = true;
            fileConvertorThread.Start();
        }
        #endregion
    }
}

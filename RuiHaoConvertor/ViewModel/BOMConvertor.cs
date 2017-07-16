using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using RuiHaoConvertor;
using System.Windows.Forms;

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
            sourceExcel = new ExcelHelper();
            exportExcel = new ExcelHelper();
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
        public void FileDispose()
        {
            try
            {
                sourceExcel.Show();
                exportExcel.Show();
                sourceExcel.Open(FilePath);
                exportExcel.Open(Environment.CurrentDirectory + @"/" + "模板.xlsx");
                //sourceExcel.Hide();
                //exportExcel.Hide();
                sourceExcel.SetActivitySheet("Export");
                exportExcel.SetActivitySheet("Export");
                for (var x = 1; x <= 10; x++)
                {
                    List<string> dataList = new List<string>();
                    //int count = -1, selectX = 0;
                    //while (count == -1)
                    //{

                    //}

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
                    for (int rowIndex = 0; rowIndex < dataList.Count; rowIndex++)
                    {
                        if (columnIndex == 11)
                            exportExcel.SetCellValue(rowIndex + 7, columnIndex, Convert.ToDouble(dataList[rowIndex]));
                        else
                            exportExcel.SetCellValue(rowIndex + 7, columnIndex, dataList[rowIndex]);
                    }
                }
                exportExcel.Save("Export");
                exportExcel.Saved();
                sourceExcel.Saved();
                exportExcel.Dispose();
                sourceExcel.Dispose();
            }
            catch (Exception e)
            {
                Message += "转换失败:" + e.Message + "\r\n";
                exportExcel.Saved();
                sourceExcel.Saved();
                exportExcel.Dispose();
                sourceExcel.Dispose();
            }
        }
        public int GetCoordinate(string title)
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

        #endregion
    }
}

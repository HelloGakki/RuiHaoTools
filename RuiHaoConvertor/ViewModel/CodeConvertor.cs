using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace RuiHaoConvertor.ViewModel
{
    public class CodeConvertor : INotifyPropertyChanged
    {
        #region " 接口实现 "

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region " 私有变量定义 "

        private string _message, _filePath;
        ExcelHelper _libraryExcel, _exportExcel;
        private List<string> _supplierCode, _companyCode;
        bool _onGunInput = true;

        #endregion

        #region " 构造函数 "

        public CodeConvertor()
        {
            _message = "感谢使用本软件\r\n";
            _message += "使用步骤:\r\n";
            _message += "①.点击\"Convertor\"按钮\r\n";
            _message += "②.打开本第一次使用时, 需要等待导入库资料以及自动打开Excel转换表格,";
            _message += "二次使用时, 只需等待自动打开的Excel表格\r\n";
            _message += "③.打开表格后, 随意点击单元格扫码输入, 转换后的件号会在右边一格的单元格中出现\r\n";
            _message += "④.完成件号转换时, 点击\"OK\"按钮会自动保存表格并关闭,也可自行保存表格,随后关闭本软件.\r\n";
            _message += "注意:若没有保存文件,直接关闭本软件,会一同关闭表格文件.\r\n";
        }

        #endregion

        #region " 属性定义 "

        public string Message
        {
            get
            {
                return _message;
            }

            set
            {
                _message = value;
                OnPropertyChanged("Message");
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
                OnPropertyChanged("FilePath");
            }
        }

        #endregion

        #region " 方法定义 "


        public void DelayConvertorStart()
        {
            Thread ConvertorStartThread = new Thread(new ThreadStart(ConvertorStart));
            ConvertorStartThread.IsBackground = true;
            ConvertorStartThread.Start();
        }
        /// <summary>
        /// 转换开始
        /// </summary>
        private void ConvertorStart()
        {
            try
            {
                if (_supplierCode == null || _companyCode == null)
                {
                    Message += "正在导入数据库...\r\n";
                    LibraryImport();
                    Message += "导入成功...\r\n" + "共 " + _supplierCode.Count.ToString() + " 项\r\n";

                }
                if (_exportExcel == null)
                {
                    Message += "正在打开新的Excel...\r\n";
                    _exportExcel = new ExcelHelper();
                    _exportExcel.Open(OnSheetChange);
                    _exportExcel.Show();
                    Message += "打开成功\r\n";
                }
                else
                {
                    Save();
                    Message += "正在打开新的Excel...\r\n";
                    _exportExcel = new ExcelHelper();
                    _exportExcel.Open(OnSheetChange);
                    _exportExcel.Show();
                    Message += "打开成功\r\n";
                }
            }
            catch (Exception e)
            {
                Message = e.Message;
            }
        }
        /// <summary>
        /// 表格新数据输入时的处理方法
        /// </summary>
        /// <param name="range"></param>
        private void OnSheetChange(Range range)
        {
            if (_onGunInput == false)
            {
                _onGunInput = true;
                return;
            }
            string data = range.Value.ToString();
            int rowIndex = range.Row;
            int columnIndex = range.Column;
            if (_supplierCode.Contains(data))
            {
                _onGunInput = false;
                string companyCode = _companyCode[_supplierCode.IndexOf(data)];
                _exportExcel.SetCellValue(rowIndex, columnIndex + 1, companyCode);
            }
            else
            {
                _onGunInput = false;
                _exportExcel.SetCellValue(rowIndex, columnIndex + 1, "库中未搜索到此供应商条码:" + data);
            }
        }
        /// <summary>
        /// 导入数据库
        /// </summary>
        private void LibraryImport()
        {
            try
            {
                _supplierCode = new List<string>();
                _companyCode = new List<string>();

                if (_libraryExcel == null)
                    _libraryExcel = new ExcelHelper();
                _libraryExcel.Show();
                _libraryExcel.Open(Environment.CurrentDirectory + @"/" + "Library.xlsx");
                _libraryExcel.Hide();
                // 导入供应商编码
                for (var y = 2; ; y++)
                {
                    object supplierData = _libraryExcel.GetCellValue(y, 1);
                    object companyData = _libraryExcel.GetCellValue(y, 2);
                    string supplierCode = supplierData == null ? null : supplierData.ToString();
                    string companyCode = companyData == null ? "库中无对应件号" : companyData.ToString();
                    if (supplierCode != null)
                    {
                        if (!_supplierCode.Contains(supplierCode))
                        {
                            _supplierCode.Add(supplierCode);
                            _companyCode.Add(companyCode);
                        }
                    }
                    else
                        break;
                }
                _libraryExcel.Close();
            }
            catch (Exception e)
            {
                Message = e.Message;
                _libraryExcel.Close();
            }
        }

        public void Save()
        {
            try
            {
                if (_exportExcel != null)
                {
                    Message += "正在保存...\r\n";
                    _exportExcel.Save("供应商转换件号");
                    _exportExcel.Saved();
                    _exportExcel.Close();
                    Message += "保存成功\r\n";
                }
            }
            catch
            {
                Message += "保存失败\r\n";
            }
            
        }

        public void Dispose()
        {
            if (_exportExcel != null)
                _exportExcel.Dispose();
            if (_libraryExcel != null)
                _libraryExcel.Dispose();
        }

        #endregion
    }
}

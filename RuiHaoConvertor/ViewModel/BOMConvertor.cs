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

        private string _message,_filePath;
        ExcelHelper excelHalper;
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
            excelHalper = new ExcelHelper();
            Message = "感谢使用本软件";
            _filePath = "";
            IsConvertor = false;
        }

        #endregion

        #region " 方法定义 "

        public void GetFilePath()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择文件";
            openFileDialog.Filter = "Excel Worksheets|*.xlsx|Excel Worksheets|*.xls";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
                FilePath = openFileDialog.FileName;
            
        }

        #endregion
    }
}

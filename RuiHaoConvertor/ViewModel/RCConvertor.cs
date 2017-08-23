using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Threading;

namespace RuiHaoConvertor.ViewModel
{
    public class RCConvertor : INotifyPropertyChanged
    {
        #region " 变更通知接口实现 "

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region " 私有变量定义 "

        private ObservableCollection<string> powerOrWithstanding, category, precision, unit;
        private string footprint, componentsValue, selectCategory, selectPowerOrWithstanding, selectPrecision, selectUnit;
        private Components components;
        private List<string> _resistanceList, _capacitanceList, _powerList, _withstandingList, _precisionList, _resUnitList, _capUnitList;
        private List<string> _codeList;
        private string message;
        ExcelHelper _exportExcel;

        #endregion

        #region " 绑定属性定义 "

        /// <summary>
        /// 功率
        /// </summary>
        public ObservableCollection<string> PowerOrWithstanding
        {
            get
            {
                return powerOrWithstanding;
            }

            set
            {
                powerOrWithstanding = value;
            }
        }

        /// <summary>
        /// 元器件子类
        /// </summary>
        public ObservableCollection<string> Category
        {
            get
            {
                return category;
            }

            set
            {
                category = value;
            }
        }
        /// <summary>
        /// 精度
        /// </summary>
        public ObservableCollection<string> Precision
        {
            get
            {
                return precision;
            }

            set
            {
                precision = value;
            }
        }
        /// <summary>
        /// 单位
        /// </summary>
        public ObservableCollection<string> Unit
        {
            get
            {
                return unit;
            }

            set
            {
                unit = value;
            }
        }
        /// <summary>
        /// 封装
        /// </summary>
        public string Footprint
        {
            get
            {
                return footprint;
            }

            set
            {
                footprint = value;
                OnPropertyChanged("Footprint");
            }
        }
        /// <summary>
        /// 元器件值
        /// </summary>
        public string ComponentsValue
        {
            get
            {
                return componentsValue;
            }

            set
            {
                this.componentsValue = value;
                OnPropertyChanged("ComponentsValue");
            }
        }
        /// <summary>
        /// 元器件父类
        /// </summary>
        public Components Components
        {
            get
            {
                return components;
            }

            set
            {
                components = value;
                OnComponentsChanged();
                OnPropertyChanged("Components");
            }
        }
        /// <summary>
        /// 选择的子类
        /// </summary>
        public string SelectCategory
        {
            get
            {
                return selectCategory;
            }

            set
            {
                selectCategory = value;
                OnPropertyChanged("SelectCategory");
            }
        }
        /// <summary>
        /// 选择的功率或者耐压值
        /// </summary>
        public string SelectPowerOrWithstanding
        {
            get
            {
                return selectPowerOrWithstanding;
            }

            set
            {
                selectPowerOrWithstanding = value;
                OnPropertyChanged("SelectPowerOrWithstanding");
            }
        }
        /// <summary>
        /// 选择的精度
        /// </summary>
        public string SelectPrecision
        {
            get
            {
                return selectPrecision;
            }

            set
            {
                selectPrecision = value;
                OnPropertyChanged("SelectPrecision");
            }
        }
        /// <summary>
        /// 选择的单位
        /// </summary>
        public string SelectUnit
        {
            get
            {
                return selectUnit;
            }

            set
            {
                selectUnit = value;
                OnPropertyChanged("SelectUnit");
            }
        }
        /// <summary>
        /// 消息
        /// </summary>
        public string Message
        {
            get
            {
                return message;
            }

            set
            {
                message = value;
                OnPropertyChanged("Message");
            }
        }

        #endregion

        #region " 构造函数 "

        public RCConvertor()
        {
            _codeList = new List<string>();
            _resistanceList = new List<string> { "贴片", "直插", "热敏", "压敏", "保险丝", "可调" };
            _capacitanceList = new List<string> { "贴片", "瓷片", "钽电容", "电解", "绦纶", "独石", "X", "Y" };
            _powerList = new List<string> { "1/16W", "1/10W", "1/8W", "1/4W", "1/2W", "1W", "2W", "3W", "5W", "其他" };
            _withstandingList = new List<string> { "6.3V", "10V", "16V", "25V", "50V", "63V", "100V", "400V", "630V", "1000V", "1600V", "2000V", "其他" };
            _precisionList = new List<string> { "0.5%", "1%", "2%", "5%", "10%", "20%", "其他" };
            _resUnitList = new List<string> { "Ω", "KΩ", "MΩ" };
            _capUnitList = new List<string> { "pF", "nF", "μF", "mF", "F" };
            Category = new ObservableCollection<string>();
            PowerOrWithstanding = new ObservableCollection<string>();
            Precision = new ObservableCollection<string>();
            Unit = new ObservableCollection<string>();
            _precisionList.ForEach(x => Precision.Add(x));
            Footprint = "0603";
            //SelectCategory = _resistanceList[0];
            //SelectPowerOrWithstanding = _powerList[2];
            //SelectPrecision = _precisionList[3];
            message = "感谢使用本软件\r\n";
            message += "使用步骤:\r\n";
            message += "①.选择电阻电容标签\r\n";
            message += "②.内容随选择标签改变后选择需要的属性, ";
            message += "填入值\r\n";
            message += "③.点击\"Confirm\"按钮, 会转换成相应公司件号显示在Message界面\r\n";
            message += "④.完成所有件号转换后, 点击\"OK\"按钮会自动保存表格,随后关闭本软件.\r\n";
            message += "注意:若没有保存文件,直接关闭本软件,会一同关闭表格文件.\r\n";
        }

        #endregion

        #region " 方法 "

        /// <summary>
        /// 元器件类别改变
        /// </summary>
        private void OnComponentsChanged()
        {
            if (Components == Components.Resistance)
            {
                Category.Clear();
                PowerOrWithstanding.Clear();
                Unit.Clear();
                _resistanceList.ForEach(x => Category.Add(x));
                _powerList.ForEach(x => PowerOrWithstanding.Add(x));
                _resUnitList.ForEach(x => Unit.Add(x));
                SelectCategory = _resistanceList[0];
                SelectPowerOrWithstanding = _powerList[2];
                SelectPrecision = _precisionList[3];
                SelectUnit = _resUnitList[0];
            }
            if (components == Components.Capacitance)
            {
                Category.Clear();
                PowerOrWithstanding.Clear();
                Unit.Clear();
                _capacitanceList.ForEach(x => Category.Add(x));
                _withstandingList.ForEach(x => PowerOrWithstanding.Add(x));
                _capUnitList.ForEach(x => Unit.Add(x));
                SelectCategory = _resistanceList[0];
                SelectPowerOrWithstanding = _withstandingList[2];
                SelectPrecision = _precisionList[3];
                SelectUnit = _capUnitList[0];
            }
        }
        /// <summary>
        /// 物料编码代号转换
        /// </summary>
        public void Coding()
        {
            string code = null;
            code = CategoryCode(SelectCategory) + Footprint + PowerOrWithstandingCode(SelectPowerOrWithstanding) + PrecisionCode(SelectPrecision)
            + ScientificCode(ComponentsValue) + "0";
            if (code != null && (code.Count() >= 13))
            {
                _codeList.Add(code);
                Message += "转换件号：" + code + "\r\n";
            }
        }
        /// <summary>
        /// 类别代号转换
        /// </summary>
        /// <param name="category"></param>
        /// <returns></returns>
        private string CategoryCode(string category)
        {
            List<string> resCode = new List<string> { "ERS", "ERT", "ERH", "ERR", "ERF", "ERV" };
            List<string> capCode = new List<string> { "ECS", "ECC", "ECD", "ECE", "ECF", "ECT", "ECX", "ECY" };

            if (Components == Components.Resistance)
            {
                if (_resistanceList.Contains(category))
                    return resCode[_resistanceList.IndexOf(category)];
                else
                    return null;
            }
            else if (Components == Components.Capacitance)
            {
                if (_capacitanceList.Contains(category))
                    return capCode[_capacitanceList.IndexOf(category)];
                else
                    return null;
            }
            else
                return null;
        }
        /// <summary>
        /// 功率代码转换
        /// </summary>
        /// <param name="power">功率值</param>
        /// <returns></returns>
        private string PowerOrWithstandingCode(string powerOrWithstanding)
        {

            if (Components == Components.Resistance)
            {
                if (_powerList.Contains(powerOrWithstanding))
                {
                    if (powerOrWithstanding != "其他")
                        return _powerList.IndexOf(powerOrWithstanding).ToString();
                    else
                        return "Z";
                }
                else
                    return null;
            }
            else if (Components == Components.Capacitance)
            {
                if (_withstandingList.Contains(powerOrWithstanding))
                {
                    if (powerOrWithstanding != "其他")
                        return Convert.ToString(_withstandingList.IndexOf(powerOrWithstanding), 16);

                    else
                        return "Z";

                }
                else
                    return null;
            }
            else
                return null;
        }
        ///// <summary>
        ///// 耐压值代号转换
        ///// </summary>
        ///// <param name="withstanding"></param>
        ///// <returns></returns>
        //private string WithstandingCode(string withstanding)
        //{

        //    if (_withstandingList.Contains(withstanding))
        //    {
        //        if (withstanding != "其他")
        //            return Convert.ToString(_withstandingList.IndexOf(withstanding), 16);

        //        else
        //            return "Z";

        //    }
        //    else
        //        return null;
        //}
        /// <summary>
        /// 精度代码转化
        /// </summary>
        /// <param name="precision">精度值</param>
        /// <returns></returns>
        private string PrecisionCode(string precision)
        {
            List<string> code = new List<string> { "D", "F", "G", "J", "K", "M", "Z" };
            if (_precisionList.Contains(precision))
                return code[_precisionList.IndexOf(precision)];
            else
                return null;
        }
        /// <summary>
        /// 科学记数转换
        /// </summary>
        /// <param name="componentsValue"></param>
        /// <returns></returns>
        private string ScientificCode(string componentsValue)
        {
            try
            {
                int zeroCount = 0;
                //int quotient = 0;
                int remainder = 0;
                int dividend = 0;

                if (Components == Components.Resistance)
                {
                    if (_resUnitList.Contains(SelectUnit))
                    {
                        if (SelectUnit == _resUnitList[0])
                        {
                            if (ComponentsValue.Contains("."))
                            {
                                if ((int)Convert.ToDouble(ComponentsValue) == 0)
                                    return ((Convert.ToDouble(ComponentsValue) * 100).ToString() + "B");
                                else
                                    return ((Convert.ToDouble(ComponentsValue) * 10).ToString() + "A");

                            }
                            //return Convert.ToUInt32(ComponentsValue) == 0 ? Convert.ToDouble(ComponentsValue) * 100).ToString() + "B" : (Convert.ToDouble(ComponentsValue) * 10).ToString() + "A";
                        }
                        zeroCount = _resUnitList.IndexOf(SelectUnit) * 3;
                    }
                }
                if (Components == Components.Capacitance)
                {
                    if (_capUnitList.Contains(SelectUnit))
                    {
                        if (SelectUnit == _capUnitList[0])
                        {
                            if (ComponentsValue.Contains("."))
                            {
                                if ((int)Convert.ToDouble(ComponentsValue) == 0)
                                    return ((Convert.ToDouble(ComponentsValue) * 100).ToString() + "B");
                                else
                                    return ((Convert.ToDouble(ComponentsValue) * 10).ToString() + "A");

                            }
                            //return Convert.ToUInt32(ComponentsValue) == 0 ? Convert.ToDouble(ComponentsValue) * 100).ToString() + "B" : (Convert.ToDouble(ComponentsValue) * 10).ToString() + "A";
                        }
                        zeroCount = _capUnitList.IndexOf(SelectUnit) * 3;
                    }
                }

                dividend = (int)(Convert.ToDouble(ComponentsValue) * Math.Pow(10, zeroCount));
                for (var i = 1; dividend >= 10; i++)
                {
                    dividend = Math.DivRem(dividend, 10, out remainder);
                    zeroCount = i;
                }
                if (zeroCount > 0)
                {
                    dividend = dividend * 10 + remainder;
                    zeroCount -= 1;
                    return dividend.ToString() + zeroCount.ToString();
                }
                else
                    return "0" + dividend.ToString() + zeroCount.ToString();
                //if (remainder > 0)
                //    dividend = dividend * 10 + remainder;
                //if (dividend >= 10)
                //    return dividend.ToString() + zeroCount.ToString();
                //else if (zeroCount > 0)
                //    return dividend.ToString() + "0" + zeroCount.ToString();
                //else
                //    return "0" + dividend.ToString() + "0";
            }
            catch (Exception e)
            {
                Message += e.Message + "\r\n";
                return null;
            }
        }
        private void SaveCode()
        {
            try
            {
                if (_codeList.Count == 0)
                {
                    Message += "导出错误，件号列表为空\r\n";
                    return;
                }
                if (_exportExcel == null)
                    _exportExcel = new ExcelHelper();
                _exportExcel.Show();
                _exportExcel.Open();
                _exportExcel.Hide();
                for (var i = 0; i < _codeList.Count; i++)
                {
                    _exportExcel.SetCellValue(i + 1, 1, i + 1);
                    _exportExcel.SetCellValue(i + 1, 2, _codeList[i]);
                }
                _exportExcel.SetCellBorder(1, 1, _codeList.Count, 2);
                _exportExcel.setCellTextByFormat(_exportExcel.GetRange(1, 1, _codeList.Count, 2), "微软雅黑", "10");
                _exportExcel.Save("CodeExoprt");
                _exportExcel.Saved();
                _exportExcel.Close();

                Message += "导出件号成功" + "\r\n" + "共" + _codeList.Count().ToString() + "项" + "\r\n";
            }
            catch (Exception e)
            {
                Message += "转换失败:" + e.Message + "\r\n";
                _exportExcel.Saved();
                _exportExcel.Close();
            }
        }

        public void DelaySaveCode()
        {
            Thread saveCodeThread = new Thread(new ThreadStart(SaveCode));
            saveCodeThread.IsBackground = true;
            saveCodeThread.Start();
        }

        public void Dispose()
        {
            if (_exportExcel != null)
                _exportExcel.Dispose();
        }

        #endregion

    }
}

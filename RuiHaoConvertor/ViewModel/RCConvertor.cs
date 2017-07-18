using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.ComponentModel;

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
        private List<string> _code;

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

        #endregion

        #region " 构造函数 "

        public RCConvertor()
        {
            _code = new List<string>();
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
            SelectCategory = _resistanceList[0];
            SelectPowerOrWithstanding = _powerList[0];
            SelectPrecision = _precisionList[0];
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
                SelectPowerOrWithstanding = _powerList[0];
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
                SelectCategory = _capacitanceList[0];
                SelectPowerOrWithstanding = _withstandingList[0];
                SelectUnit = _capUnitList[0];
            }
        }
        /// <summary>
        /// 物料编码代号转换
        /// </summary>
        public void Coding()
        {
            string code = "";
            code = CategoryCode(SelectCategory) + Footprint + PowerCode(SelectPowerOrWithstanding) + PrecisionCode(SelectPrecision);

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
                    return "";
            }
            else if (Components == Components.Capacitance)
            {
                if (_capacitanceList.Contains(category))
                    return capCode[_capacitanceList.IndexOf(category)];
                else
                    return "";
            }
            else
                return "";
        }
        /// <summary>
        /// 功率代码转换
        /// </summary>
        /// <param name="power">功率值</param>
        /// <returns></returns>
        private string PowerCode(string power)
        {
            if (_powerList.Contains(power))
            {
                if (power != "其他")
                    return _powerList.IndexOf(power).ToString();
                else
                    return "Z";
            }
            else
                return "";
        }
        /// <summary>
        /// 耐压值代号转换
        /// </summary>
        /// <param name="withstanding"></param>
        /// <returns></returns>
        private string WithstandingCode(string withstanding)
        {
           
            if (_withstandingList.Contains(withstanding))
            {
                if (withstanding != "其他")
                    return Convert.ToString(_withstandingList.IndexOf(withstanding), 16);
                
                else
                    return "Z";
                
            }
            else
                return "";
        }
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
                return "";
        }


        #endregion

    }
}

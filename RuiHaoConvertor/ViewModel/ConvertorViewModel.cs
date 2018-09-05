using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Windows.Input;

namespace RuiHaoConvertor.ViewModel
{
    public class ConvertorViewModel:INotifyPropertyChanged
    {
        private BOMConvertor bomConvertor;
        private RCConvertor rcConvertor;
        private CodeConvertor codeConvertor;

        public ConvertorViewModel()
        {
            BomConvertor = new BOMConvertor();
            RCConvertor = new RCConvertor();
            CodeConvertor = new CodeConvertor();
        }

        public BOMConvertor BomConvertor
        {
            get
            {
                return bomConvertor;
            }

            set
            {
                bomConvertor = value;
                OnPropertyChanged("BomConvertor");
            }
        }

        public RCConvertor RCConvertor
        {
            get
            {
                return rcConvertor;
            }

            set
            {
                rcConvertor = value;
                OnPropertyChanged("RCConvertor");
            }
        }

        public CodeConvertor CodeConvertor
        {
            get
            {
                return codeConvertor;
            }

            set
            {
                codeConvertor = value;
                OnPropertyChanged("CodeConvertor");
            }
        }

        #region " 变更通知接口实现 "

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            //  PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region " 命令定义 "

        // 打开About窗口
        public static RoutedCommand ShowAboutCommand =
            new RoutedCommand("Show About", typeof(ConvertorViewModel),
                new InputGestureCollection(new List<InputGesture> { new KeyGesture(Key.F1) }));

        #endregion
    }
}

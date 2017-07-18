using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace RuiHaoConvertor.ViewModel
{
    public class ConvertorViewModel:INotifyPropertyChanged
    {
        private BOMConvertor bomConvertor;
        private RCConvertor rcConvertor;

        public ConvertorViewModel()
        {
            BomConvertor = new BOMConvertor();
            RCConvertor = new RCConvertor();
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

        #region " 变更通知接口实现 "

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}

using System.Collections.Generic;
using System.ComponentModel;

namespace CreateUsers.Model
{
    public class Cities : INotifyPropertyChanged
    {
        private string _displayName;
        private string _name;
        private string _adress;
        private List<string> _groups;

        public List<string> Groups
        {
            get { return _groups; }
            set 
            { 
                _groups = value;
                OnPropertyChanged("Groups");
            }
        }

        public string Adress
        {
            get { return _adress; }
            set 
            { 
                _adress = value;
                OnPropertyChanged("Adress");
            }
        }

        public string Name
        {
            get { return _name; }
            set 
            { 
                _name = value;
                OnPropertyChanged("Name");
            }
        }

        public string DisplayName
        {
            get { return _displayName; }
            set 
            { 
                _displayName = value;
                OnPropertyChanged("DisplayName");
            }
        }


        #region реализация INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion
    }
}

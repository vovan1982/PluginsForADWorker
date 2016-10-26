using System.Collections.Generic;
using System.ComponentModel;

namespace GetInfoFromExchange.Model
{
    public class UserExchangeInfo : INotifyPropertyChanged
    {
        private string _user;
        private bool _isEnableActiveSync;
        private bool _isEnableWebApp;
        private bool _isEnableMAPI;
        private bool _isEnablePOP;
        private bool _isEnableIMAP;
        private string _forwardingAddress;
        private string _forwardingSmtpAddress;
        private string _policyActiveSync;
        private string _policyWebApp;
        private string _emailAddress;
        private string _primaryEmailAddres;
        private List<InboxMessageRuleInfo> _indoxMessageRule;
        private UserAutoReplayInfo _autoReplay;
        private List<MobileDeviceInfo> _mobileDevice;

        public List<MobileDeviceInfo> MobileDevice
        {
            get { return _mobileDevice; }
            set 
            { 
                _mobileDevice = value;
                OnPropertyChanged("MobileDevice");
            }
        }

        public UserAutoReplayInfo AutoReplay
        {
            get { return _autoReplay; }
            set 
            { 
                _autoReplay = value;
                OnPropertyChanged("AutoReplay");
            }
        }

        public List<InboxMessageRuleInfo> IndoxMessageRule
        {
            get { return _indoxMessageRule; }
            set 
            { 
                _indoxMessageRule = value;
                OnPropertyChanged("IndoxMessageRule");
            }
        }


        public string PrimaryEmailAddres
        {
            get { return _primaryEmailAddres; }
            set 
            { 
                _primaryEmailAddres = value;
                OnPropertyChanged("PrimaryEmailAddres");
            }
        }

        public string EmailAddress
        {
            get { return _emailAddress; }
            set 
            { 
                _emailAddress = value;
                OnPropertyChanged("EmailAddress");
            }

        }

        public string PolicyWebApp
        {
            get { return _policyWebApp; }
            set 
            { 
                _policyWebApp = value;
                OnPropertyChanged("PolicyWebApp");
            }
        }

        public string PolicyActiveSync
        {
            get { return _policyActiveSync; }
            set 
            { 
                _policyActiveSync = value;
                OnPropertyChanged("PolicyActiveSync");
            }
        }

        public string ForwardingSmtpAddress
        {
            get { return _forwardingSmtpAddress; }
            set 
            { 
                _forwardingSmtpAddress = value;
                OnPropertyChanged("ForwardingSmtpAddress");
            }
        }

        public string ForwardingAddress
        {
            get { return _forwardingAddress; }
            set 
            { 
                _forwardingAddress = value;
                OnPropertyChanged("ForwardingAddress");
            }
        }

        public bool IsEnableIMAP
        {
            get { return _isEnableIMAP; }
            set 
            { 
                _isEnableIMAP = value;
                OnPropertyChanged("IsEnableIMAP");
            }
        }

        public bool IsEnablePOP
        {
            get { return _isEnablePOP; }
            set 
            { 
                _isEnablePOP = value;
                OnPropertyChanged("IsEnablePOP");
            }
        }

        public bool IsEnableMAPI
        {
            get { return _isEnableMAPI; }
            set 
            { 
                _isEnableMAPI = value;
                OnPropertyChanged("IsEnableMAPI");
            }
        }

        public bool IsEnableWebApp
        {
            get { return _isEnableWebApp; }
            set 
            { 
                _isEnableWebApp = value;
                OnPropertyChanged("IsEnableWebApp");
            }
        }

        public bool IsEnableActiveSync
        {
            get { return _isEnableActiveSync; }
            set 
            { 
                _isEnableActiveSync = value;
                OnPropertyChanged("IsEnableActiveSync");
            }
        }

        public string User
        {
            get { return _user; }
            set 
            { 
                _user = value;
                OnPropertyChanged("User");
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

using System;
using System.ComponentModel;

namespace RussificationOutlookFolders.Model
{
    public class User : INotifyPropertyChanged
    {
        private string _placeInAD;
        private string _name;
        private string _surname;
        private string _login;
        private string _nameInAD;
        private string _post;
        private string _department;
        private string _city;
        private string _mail;
        private string _phoneInt;
        private string _phoneMob;
        private string _displayName;
        private string _organization;
        private string _adress;
        private DateTime? _accountExpireDate;
        private DateTime _passExpireDate;
        private bool _passMustBeChange;
        private bool _accountIsDisable;
        private bool _accountIsLock;

        public User()
        {

        }

        public bool AccountIsLock
        {
            get { return _accountIsLock; }
            set
            {
                _accountIsLock = value;
                OnPropertyChanged("AccountIsLock");
            }
        }

        public string NameInAD
        {
            get { return _nameInAD; }
            set
            {
                _nameInAD = value;
                OnPropertyChanged("NameInAD");
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

        public string Organization
        {
            get { return _organization; }
            set
            {
                _organization = value;
                OnPropertyChanged("Organization");
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

        public string Login
        {
            get { return _login; }
            set
            {
                _login = value;
                OnPropertyChanged("Login");
            }
        }

        public bool AccountIsDisable
        {
            get { return _accountIsDisable; }
            set
            {
                _accountIsDisable = value;
                OnPropertyChanged("AccountIsDisable");
            }
        }

        public bool PassMustBeChange
        {
            get { return _passMustBeChange; }
            set
            {
                _passMustBeChange = value;
                OnPropertyChanged("PassMustBeChange");
            }
        }

        public DateTime PassExpireDate
        {
            get { return _passExpireDate; }
            set
            {
                _passExpireDate = value;
                OnPropertyChanged("PassExpireDate");
            }
        }

        public DateTime? AccountExpireDate
        {
            get { return _accountExpireDate; }
            set
            {
                _accountExpireDate = value;
                OnPropertyChanged("AccountExpireDate");
            }
        }

        public string PhoneMob
        {
            get { return _phoneMob; }
            set
            {
                _phoneMob = value;
                OnPropertyChanged("PhoneMob");
            }
        }

        public string PhoneInt
        {
            get { return _phoneInt; }
            set
            {
                _phoneInt = value;
                OnPropertyChanged("PhoneInt");
            }
        }

        public string Mail
        {
            get { return _mail; }
            set
            {
                _mail = value;
                OnPropertyChanged("Mail");
            }
        }

        public string City
        {
            get { return _city; }
            set
            {
                _city = value;
                OnPropertyChanged("City");
            }
        }

        public string Department
        {
            get { return _department; }
            set
            {
                _department = value;
                OnPropertyChanged("Department");
            }
        }

        public string Post
        {
            get { return _post; }
            set
            {
                _post = value;
                OnPropertyChanged("Post");
            }
        }

        public string Surname
        {
            get { return _surname; }
            set
            {
                _surname = value;
                OnPropertyChanged("Surname");
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

        public string PlaceInAD
        {
            get { return _placeInAD; }
            set
            {
                _placeInAD = value;
                OnPropertyChanged("PlaceInAD");
            }
        }

        public void copyData(User userForCopy)
        {
            AccountIsLock = userForCopy.AccountIsLock;
            NameInAD = userForCopy.NameInAD;
            DisplayName = userForCopy.DisplayName;
            Organization = userForCopy.Organization;
            Login = userForCopy.Login;
            AccountIsDisable = userForCopy.AccountIsDisable;
            PassMustBeChange = userForCopy.PassMustBeChange;
            PassExpireDate = userForCopy.PassExpireDate;
            AccountExpireDate = userForCopy.AccountExpireDate;
            PhoneMob = userForCopy.PhoneMob;
            PhoneInt = userForCopy.PhoneInt;
            Mail = userForCopy.Mail;
            Adress = userForCopy.Adress;
            City = userForCopy.City;
            Department = userForCopy.Department;
            Post = userForCopy.Post;
            Surname = userForCopy.Surname;
            Name = userForCopy.Name;
            PlaceInAD = userForCopy.PlaceInAD;
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

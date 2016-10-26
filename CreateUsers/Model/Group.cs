using System.Collections.Generic;
using System.ComponentModel;

namespace CreateUsers.Model
{
    public class Group : INotifyPropertyChanged
    {
        private string _placeInAD;
        private string _name;
        private string _displayName;
        private string _nameWin;
        private string _description;
        private int _type;
        private string _mail;
        private List<string> _users;
        private List<string> _inGroups;

        public Group()
        {

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

        public List<string> InGroups
        {
            get { return _inGroups; }
            set
            {
                _inGroups = value;
                OnPropertyChanged("InGroups");
            }
        }

        public List<string> Users
        {
            get { return _users; }
            set
            {
                _users = value;
                OnPropertyChanged("Users");
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

        public int Type
        {
            get { return _type; }
            set
            {
                _type = value;
                OnPropertyChanged("Type");
            }
        }

        public string Description
        {
            get { return _description; }
            set
            {
                _description = value;
                OnPropertyChanged("Description");
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

        public string NameWin
        {
            get { return _nameWin; }
            set
            {
                _nameWin = value;
                OnPropertyChanged("NameWin");
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

        public void copyData(Group groupForCopy)
        {
            Name = groupForCopy.Name;
            NameWin = groupForCopy.NameWin;
            DisplayName = groupForCopy.DisplayName;
            PlaceInAD = groupForCopy.PlaceInAD;
            Description = groupForCopy.Description;
            Mail = groupForCopy.Mail;
            Type = groupForCopy.Type;
            Users = groupForCopy.Users;
            InGroups = groupForCopy.InGroups;
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

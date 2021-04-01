using System.Collections.Generic;

namespace CreateUsersPPK.Model
{
    public class NewUser
    {
        private string _name;
        private string _surname;
        private string _placeInAD;
        private string _login;
        private string _nameInAD;
        private string _city;
        private string _organization;
        private string _adress;
        private string _telephoneNumber;
        private string _mobile;
        private string _post;
        private string _department;
        private string _mail;
        private string _mailDataBase;
        private bool _activeSync;
        private string _activeSyncPolicy;
        private bool _owa;
        private string _owaPolicy;
        private List<string> _groups;

        public NewUser()
        {
            _groups = new List<string>();
        }

        public List<string> Groups
        {
            get { return _groups; }
            set { _groups = value; }
        }

        public string OwaPolicy
        {
            get { return _owaPolicy; }
            set { _owaPolicy = value; }
        }

        public bool Owa
        {
            get { return _owa; }
            set { _owa = value; }
        }

        public string ActiveSyncPolicy
        {
            get { return _activeSyncPolicy; }
            set { _activeSyncPolicy = value; }
        }

        public bool ActiveSync
        {
            get { return _activeSync; }
            set { _activeSync = value; }
        }

        public string MailDataBase
        {
            get { return _mailDataBase; }
            set { _mailDataBase = value; }
        }

        public string Mail
        {
            get { return _mail; }
            set { _mail = value; }
        }

        public string Department
        {
            get { return _department; }
            set { _department = value; }
        }

        public string Post
        {
            get { return _post; }
            set { _post = value; }
        }

        public string Adress
        {
            get { return _adress; }
            set { _adress = value; }
        }

        public string TelephoneNumber
        {
            get { return _telephoneNumber; }
            set { _telephoneNumber = value; }
        }

        public string Mobile
        {
            get { return _mobile; }
            set { _mobile = value; }
        }

        public string Organization
        {
            get { return _organization; }
            set { _organization = value; }
        }

        public string City
        {
            get { return _city; }
            set { _city = value; }
        }

        public string NameInAD
        {
            get { return _nameInAD; }
            set { _nameInAD = value; }
        }

        public string Login
        {
            get { return _login; }
            set { _login = value; }
        }

        public string PlaceInAD
        {
            get { return _placeInAD; }
            set { _placeInAD = value; }
        }

        public string Surname
        {
            get { return _surname; }
            set { _surname = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
    }
}

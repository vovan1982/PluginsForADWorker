using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ChangeSurname
{
    public class ProcessRunData : INotifyPropertyChanged
    {
        private ImageSource _image;
        private string _text;

        public ImageSource Image
        {
            get { return _image; }
            set
            {
                _image = value;
                OnPropertyChanged("Image");
            }
        }
        public string Text
        {
            get { return _text; }
            set
            {
                _text = value;
                OnPropertyChanged("Text");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
    /// <summary>
    /// Логика взаимодействия для ControlPlMain.xaml
    /// </summary>
    public partial class ControlPlMain : UserControl
    {
        private string _login; // Логин для соединения с почтовым сервером
        private string _pathToConfig; // Полный путь к конфигурационному файлу плагина
        private SecureString _password; // Пароль для соединения с почтовым сервером
        private DirectoryEntry _ADSession; // Сессия с АД
        private PrincipalContext _principalContext; // Контекст с АД
        private Dictionary<string, string> configData; // Данные конфигурационного файла
        private ObservableCollection<ProcessRunData> _processRun; // Лог
        private bool _userIsSelected; // Признак того что пользователь выбран

        // Конструктор
        public ControlPlMain(IPlugin plug, string login, string pass, string pathToConfig, DirectoryEntry entry, PrincipalContext context)
        {
            InitializeComponent();
            _login = login;
            _password = new SecureString();
            foreach (char x in pass)
            {
                _password.AppendChar(x);
            }
            _ADSession = entry;
            _principalContext = context;
            _pathToConfig = pathToConfig;
            _userIsSelected = false;
            _processRun = new ObservableCollection<ProcessRunData>();
            log.ItemsSource = _processRun;
            lbVersion.Content = lbVersion.Content + " " + plug.Version;
        }
        
        // Нажата кнопка выбора сотрудника
        private void btSelectUser_Click(object sender, RoutedEventArgs e)
        {
            SelectUser _dwSU = null;
            if (sender is string)
            {
                if ((string)sender == "dnUserForChange")
                {
                    _dwSU = new SelectUser(_ADSession, _principalContext, dnUserForChange.Text);
                    _dwSU.ShowDialog();
                }
            }
            else
            {
                _dwSU = new SelectUser(_ADSession, _principalContext);
                _dwSU.ShowDialog();
            }
            if (_dwSU.DialogResult == true)
            {
                dnUserForChange.Text = _dwSU.SelectedUser;
                _userIsSelected = true;
                dnUserForChange.FontWeight = FontWeights.Bold;
                dnUserForChange.TextDecorations = TextDecorations.Underline;
            }
        }
        // Нажата кнопка изменения фамилии
        private void btChangeSurname_Click(object sender, RoutedEventArgs e)
        {
            if (!_userIsSelected)
            {
                MessageBox.Show("Не выбран пользователь!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(newSurname.Text))
            {
                MessageBox.Show("Не указана новая фамилия сотрудника!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(_pathToConfig))
            {
                MessageBox.Show("Отсутствует конфигурационный файл плагина!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string errorMesg = "";
            string dnUserToChangeSurname = dnUserForChange.Text;
            string newUserSurname = newSurname.Text.Trim();

            Thread t = new Thread(() =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Clear();
                    _processRun.Add(new ProcessRunData { Text = "---- Изменение фамилии сотрудника " + dnUserToChangeSurname.Split(',')[0].Substring(3) + "на " + newUserSurname + " ----" });
                    btSelectUser.IsEnabled = false;
                    btChangeSurname.IsEnabled = false;
                }));

                #region Чтение конфигурационного файла плагина
                ProcessRunData itemProcess = new ProcessRunData { Text = "Чтение конфигурационного файла плагина..." };
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Чтение конфигурационного файла плагина..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                configData = GetConfigData(_pathToConfig, ref errorMesg);
                if (!string.IsNullOrWhiteSpace(errorMesg))
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка чтения конфигурационного файла:\r\n" + errorMesg;
                        btSelectUser.IsEnabled = true;
                        btChangeSurname.IsEnabled = true;
                    }));
                    return;
                }
                if (!configData.ContainsKey("mailserver") || !configData.ContainsKey("maildomain") || !configData.ContainsKey("mailprimarydomain") || !configData.ContainsKey("lyncserver") || !configData.ContainsKey("lyncpool"))
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Отсутсвуют требуемые параметры в конфигурационном файле!!!";
                        btSelectUser.IsEnabled = true;
                        btChangeSurname.IsEnabled = true;
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Чтение конфигурационного файла плагина";
                }));
                #endregion
                
                //============================================================
                DirectorySearcher dirSearcher = new DirectorySearcher(_ADSession);
                dirSearcher.SearchScope = SearchScope.Subtree;
                dirSearcher.PropertiesToLoad.Add("cn");
                string newUserLogin = "";
                UserPrincipal userToModify = UserPrincipal.FindByIdentity(_principalContext, IdentityType.DistinguishedName, dnUserToChangeSurname);
                if (userToModify == null)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        _processRun.Add(new ProcessRunData { Text = "Не удалось получить информацию по текущему пользователю в АД!!!", 
                            Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute)) });
                        btSelectUser.IsEnabled = true;
                        btChangeSurname.IsEnabled = true;
                    }));
                    return;
                }
                //============================================================

                #region Проверка нового логина в домене
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Проверка нового логина пользователя " + dnUserToChangeSurname.Split(',')[0].Substring(3) + " в домене..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));

                newUserLogin = translit(userToModify.GivenName + "." + newUserSurname);
                if (newUserLogin.Length > 20)
                    newUserLogin = translit(userToModify.GivenName.Substring(0, 1) + "." + newUserSurname);
                if (newUserLogin.Length > 20)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка:\r\nНе удалось сократить логин пользователя до 20 символов!!!";
                        btSelectUser.IsEnabled = true;
                        btChangeSurname.IsEnabled = true;
                    }));
                    return;
                }
                for (int i = 1; i <= userToModify.GivenName.Length; i++)
                {
                    dirSearcher.Filter = string.Format("(&(objectClass=user)(sAMAccountName={0}))", newUserLogin);
                    SearchResult searchResults = dirSearcher.FindOne();
                    if (searchResults != null)
                    {
                        newUserLogin = translit(userToModify.GivenName.Substring(0, i) + "." + newUserSurname);
                        if (newUserLogin.Length > 20)
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text = "Ошибка:\r\nНе удалось сократить логин пользователя до 20 символов!!!";
                                btSelectUser.IsEnabled = true;
                                btChangeSurname.IsEnabled = true;
                            }));
                            return;
                        }
                    }
                    else
                    {
                        break;
                    }
                    if (i == userToModify.GivenName.Length)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Ошибка:\r\nНе удалось сформировать новый логин пользователя!!!";
                            btSelectUser.IsEnabled = true;
                            btChangeSurname.IsEnabled = true;
                        }));
                        return;
                    }
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Проверка нового логина пользователя " + dnUserToChangeSurname.Split(',')[0].Substring(3) + " в домене";
                }));
                #endregion

                #region Внесение изменений в учетную запись пользователя в АД
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Изменение фамилии пользователя " + dnUserToChangeSurname.Split(',')[0].Substring(3) + " в АД..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    if (userToModify.Surname != newUserSurname)
                    {
                        userToModify.DisplayName = translit(newUserSurname + " " + userToModify.GivenName);
                        userToModify.Surname = newUserSurname;
                        userToModify.SamAccountName = newUserLogin;
                        string[] upn = userToModify.UserPrincipalName.Split('@');
                        userToModify.UserPrincipalName = newUserLogin + "@" + upn[1];
                        userToModify.Save();
                        dirSearcher.Filter = string.Format("(&(objectClass=user)(sAMAccountName={0}))", userToModify.SamAccountName);
                        dirSearcher.PropertiesToLoad.Add("name");
                        SearchResult searchResults = dirSearcher.FindOne();
                        if (searchResults != null)
                        {
                            DirectoryEntry entryToUpdate = searchResults.GetDirectoryEntry();
                            if (entryToUpdate != null)
                            {
                                entryToUpdate.Rename("CN=" + translit(userToModify.GivenName + " " + newUserSurname));
                                entryToUpdate.CommitChanges();
                            }
                            else
                            {
                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                    itemProcess.Text = "Ошибка внесения изменений в учетную запись пользователя в АД:\r\nНе удалось получить запись пользователя в АД!!!";
                                    btSelectUser.IsEnabled = true;
                                    btChangeSurname.IsEnabled = true;
                                }));
                                return;
                            }
                        }
                        else
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text = "Ошибка внесения изменений в учетную запись пользователя в АД:\r\nНе удалось найти пользователя в АД!!!";
                                btSelectUser.IsEnabled = true;
                                btChangeSurname.IsEnabled = true;
                            }));
                            return;
                        }
                        Thread.Sleep(15000);
                    }
                }
                catch (Exception exc)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка внесения изменений в учетную запись пользователя в АД:\r\n" + exc.Message;
                        btSelectUser.IsEnabled = true;
                        btChangeSurname.IsEnabled = true;
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Изменение фамилии пользователя " + dnUserToChangeSurname.Split(',')[0].Substring(3) + " в АД";
                }));
                #endregion

                //============================================================
                PSCredential credential = new PSCredential(_login, _password);
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo((new Uri(configData["mailserver"])), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
                connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;
                Runspace runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(connectionInfo);
                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                string[] arrMailDomains = configData["maildomain"].Split(',');
                string newMailUserLogin = "";
                string[] userLoginArr = newUserLogin.Split('.');
                //============================================================
                
                #region Внесение изменений в Exchange

                #region Подключение к Exchange Server
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Подключение к Exchange Server..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    runspace.Open();
                    powershell.Runspace = runspace;
                }
                catch (Exception exp)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка подключения к Exchange Server:\r\n" + exp.Message;
                        btSelectUser.IsEnabled = true;
                        btChangeSurname.IsEnabled = true;
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Подключение к Exchange Server";
                }));
                #endregion

                #region Проверка нового электронного адреса
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Проверка нового электронного адреса пользователя " + dnUserToChangeSurname.Split(',')[0].Substring(3) + "..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    for (int i = 1; i <= userLoginArr[0].Length; i++)
                    {
                        newMailUserLogin = userLoginArr[0].Substring(0, i) + userLoginArr[1];
                        command.AddCommand("Get-Mailbox");
                        command.AddParameter("Identity", newMailUserLogin + "@" + configData["mailprimarydomain"].Trim());
                        command.AddCommand("Select-Object");
                        command.AddParameter("Property", "SamAccountName");
                        powershell.Commands = command;
                        var result = powershell.Invoke();
                        command.Clear();
                        if (result.Count <= 0)
                        {
                            powershell.Streams.Error.Clear();
                            break;
                        }
                        else
                        {
                            if (result[0].Properties["SamAccountName"].Value.ToString() == userToModify.SamAccountName)
                            {
                                powershell.Streams.Error.Clear();
                                break;
                            }
                        }
                        if (i == userLoginArr[0].Length)
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text = "Ошибка:\r\nНе удалось сформировать эл.адресс пользователя!!!";
                            }));
                            runspace.Dispose();
                            runspace = null;
                            powershell.Dispose();
                            powershell = null;
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                btSelectUser.IsEnabled = true;
                                btChangeSurname.IsEnabled = true;
                            }));
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка проверки нового электронного адреса пользователя:\r\n" + ex.Message;
                    }));
                    command.Clear();
                    runspace.Dispose();
                    runspace = null;
                    powershell.Dispose();
                    powershell = null;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        btSelectUser.IsEnabled = true;
                        btChangeSurname.IsEnabled = true;
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Проверка нового электронного адреса пользователя " + dnUserToChangeSurname.Split(',')[0].Substring(3);
                }));
                #endregion

                #region Изменение псевдонима пользователя в Exchange
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Изменение псевдонима пользователя " + dnUserToChangeSurname.Split(',')[0].Substring(3) + " в Exchange..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    command.AddCommand("Set-Mailbox");
                    command.AddParameter("Identity", userToModify.SamAccountName);
                    command.AddParameter("Alias", userToModify.SamAccountName);
                    powershell.Commands = command;
                    powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Text = "";
                        }));
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка изменения псевдонима пользователя в Exchange:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text += errorMsg;
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            btSelectUser.IsEnabled = true;
                            btChangeSurname.IsEnabled = true;
                        }));
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка изменения псевдонима пользователя в Exchange:\r\n" + ex.Message;
                    }));
                    command.Clear();
                    runspace.Dispose();
                    runspace = null;
                    powershell.Dispose();
                    powershell = null;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        btSelectUser.IsEnabled = true;
                        btChangeSurname.IsEnabled = true;
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Изменение псевдонима пользователя " + dnUserToChangeSurname.Split(',')[0].Substring(3) + "в Exchange";
                }));
                #endregion

                #region Добавление эл.адресов, установка основного адреса пользователю в Exchange
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Добавление эл.адресов, установка основного адреса пользователю " + dnUserToChangeSurname.Split(',')[0].Substring(3) + " в Exchange..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    string[] addEmailString = new string[] {};
                    command.AddCommand("Get-Mailbox");
                    command.AddParameter("Identity", userToModify.SamAccountName);
                    command.AddCommand("Select-Object");
                    command.AddParameter("ExpandProperty", "EmailAddresses");
                    powershell.Commands = command;
                    var result = powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Text = "";
                        }));
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка добавления эл.адресов, установка основного адреса пользователю:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text += errorMsg;
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            btSelectUser.IsEnabled = true;
                            btChangeSurname.IsEnabled = true;
                        }));
                        return;
                    }
                    if (result.Count > 0)
                    {
                        for (int i = 0; i < result.Count; i++)
                        {
                            var cmd = result[i];
                            string cmdletName = cmd.Properties["ProxyAddressString"].Value.ToString();
                            if (cmdletName.StartsWith("SMTP"))
                            {
                                Array.Resize(ref addEmailString, addEmailString.Length + 1);
                                addEmailString[addEmailString.Length - 1] = "smtp:" + cmdletName.Split(':')[1];
                            }
                            else
                            {
                                Array.Resize(ref addEmailString, addEmailString.Length + 1);
                                addEmailString[addEmailString.Length - 1] = cmdletName;
                            }
                        }
                    }
                    command.AddCommand("Set-Mailbox");
                    command.AddParameter("Identity", userToModify.SamAccountName);
                    for (int i = 0; i < arrMailDomains.Length; i++)
                    {
                        int index = 0;
                        if (!containsInArr(addEmailString, "smtp:" + newMailUserLogin + "@" + arrMailDomains[i], ref index))
                        {
                            Array.Resize(ref addEmailString, addEmailString.Length + 1);
                            if (newMailUserLogin + "@" + arrMailDomains[i] == newMailUserLogin + "@" + configData["mailprimarydomain"].Trim())
                            {
                                addEmailString[addEmailString.Length - 1] = "SMTP:" + newMailUserLogin + "@" + arrMailDomains[i];
                            }
                            else
                            {
                                addEmailString[addEmailString.Length - 1] = "smtp:" + newMailUserLogin + "@" + arrMailDomains[i];
                            }
                        }
                        else
                        {
                            if (newMailUserLogin + "@" + arrMailDomains[i] == newMailUserLogin + "@" + configData["mailprimarydomain"].Trim())
                            {
                                addEmailString[index] = "SMTP:" + newMailUserLogin + "@" + arrMailDomains[i];
                            }
                        }
                    }
                    command.AddParameter("EmailAddresses", addEmailString);
                    powershell.Commands = command;
                    powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Text = "";
                        }));
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка добавления эл.адресов, установка основного адреса пользователю:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text += errorMsg;
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            btSelectUser.IsEnabled = true;
                            btChangeSurname.IsEnabled = true;
                        }));
                        return;
                    }
                    Thread.Sleep(15000);
                }
                catch (Exception ex)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка добавления эл.адресов, установка основного адреса пользователю:\r\n" + ex.Message;
                    }));
                    command.Clear();
                    runspace.Dispose();
                    runspace = null;
                    powershell.Dispose();
                    powershell = null;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        btSelectUser.IsEnabled = true;
                        btChangeSurname.IsEnabled = true;
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Добавление эл.адресов пользователю, установка основного адреса " + dnUserToChangeSurname.Split(',')[0].Substring(3) + " в Exchange";
                }));
                #endregion

                #endregion

                //============================================================
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
                command.Clear();
                connectionInfo = new WSManConnectionInfo((new Uri(configData["lyncserver"])), "http://schemas.microsoft.com/powershell/Microsoft.PowerShell", credential);
                connectionInfo.AuthenticationMechanism = AuthenticationMechanism.NegotiateWithImplicitCredential;
                connectionInfo.SkipCACheck = true;
                connectionInfo.SkipCNCheck = true;
                connectionInfo.SkipRevocationCheck = true;
                runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(connectionInfo);
                powershell = PowerShell.Create();
                //============================================================

                #region Внесение изменений в Lync

                #region Подключение к Lync Server
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Подключение к Lync Server..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    runspace.Open();
                    powershell.Runspace = runspace;
                }
                catch (Exception exp)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка подключения к Lync Server:\r\n" + exp.Message;
                        btSelectUser.IsEnabled = true;
                        btChangeSurname.IsEnabled = true;
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Подключение к Lync Server";
                }));
                #endregion

                #region Внесение изменений в Lync
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Внесение изменений в Lync для пользователя " + dnUserToChangeSurname.Split(',')[0].Substring(3) + "..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    command.AddCommand("enable-csuser");
                    command.AddParameter("Identity", userToModify.SamAccountName);
                    command.AddParameter("registrarPool", configData["lyncpool"]);
                    command.AddParameter("SipAddressType", "EmailAddress");
                    powershell.Commands = command;
                    powershell.Invoke();
                    command.Clear();
                }
                catch (Exception ex)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка внесения изменений в Lync:\r\n" + ex.Message;
                    }));
                    runspace.Dispose();
                    runspace = null;
                    powershell.Dispose();
                    powershell = null;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        btSelectUser.IsEnabled = true;
                        btChangeSurname.IsEnabled = true;
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/ChangeSurname;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Внесение изменений в Lync для пользователя " + dnUserToChangeSurname.Split(',')[0].Substring(3);
                }));
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
                #endregion

                #endregion

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "---- Процесс завершён! ----" });
                    dnUserForChange.Text = "";
                    dnUserForChange.FontWeight = FontWeights.Normal;
                    dnUserForChange.TextDecorations = null;
                    newSurname.Text = "";
                    btSelectUser.IsEnabled = true;
                    btChangeSurname.IsEnabled = true;
                }));
            });
            t.IsBackground = true;
            t.Start();
        }
        // Чтение конфигурационного файла
        private Dictionary<string, string> GetConfigData(string configFilePach, ref string errorStr)
        {
            Dictionary<string, string> configData = new Dictionary<string, string>();
            string configDataFull = "";

            try
            {
                using (StreamReader sr = new StreamReader(configFilePach))
                {
                    configDataFull = sr.ReadToEnd();
                    string[] configArr = configDataFull.Split('\n');
                    for (int i = 0; i < configArr.Length; i++)
                    {
                        if (!string.IsNullOrWhiteSpace(configArr[i]))
                        {
                            string[] conf = configArr[i].Split('=');
                            configData.Add(conf[0], conf[1].TrimEnd('\r'));
                        }
                    }
                }
            }
            catch (Exception e)
            {
                errorStr += "Ошибка чтения файла конфигурации " + configFilePach + ":\r\n" + e.Message;
                configData = null;
            }
            return configData;
        }
        // Двойное нажание на элемент лога
        private void logItem_MouseDoubleClick(object sender, RoutedEventArgs e)
        {
            ShowLogText _dwSLT = new ShowLogText(((ProcessRunData)log.SelectedItem).Text);
            _dwSLT.ShowDialog();
        }
        // Транслитерация
        private string translit(string inputText)
        {
            Dictionary<char, string> translitDictionary = new Dictionary<char, string>() 
            { 
                {'а', "a"}, {'А', "A"}, {'б', "b"}, {'Б', "B"},
                {'в', "v"}, {'В', "V"}, {'г', "g"}, {'Г', "G"},
                {'д', "d"}, {'Д', "D"}, {'е', "e"}, {'Е', "E"},
                {'ё', "e"}, {'Ё', "E"}, {'ж', "zh"}, {'Ж', "Zh"},
                {'з', "z"}, {'З', "Z"}, {'и', "i"}, {'И', "I"},
                {'й', "y"}, {'Й', "Y"}, {'к', "k"}, {'К', "K"},
                {'л', "l"}, {'Л', "L"}, {'м', "m"}, {'М', "M"},
                {'н', "n"}, {'Н', "N"}, {'о', "o"}, {'О', "O"},
                {'п', "p"}, {'П', "P"}, {'р', "r"}, {'Р', "R"},
                {'с', "s"}, {'С', "S"}, {'т', "t"}, {'Т', "T"},
                {'у', "u"}, {'У', "U"}, {'ф', "f"}, {'Ф', "F"},
                {'х', "kh"}, {'Х', "Kh"}, {'ц', "ts"}, {'Ц', "Ts"},
                {'ч', "ch"}, {'Ч', "Ch"}, {'ш', "sh"}, {'Ш', "Sh"},
                {'щ', "shch"}, {'Щ', "Shch"}, {'ъ', ""}, {'Ъ', ""},
                {'ы', "y"}, {'Ы', "Y"}, {'ь', ""}, {'Ь', ""},
                {'э', "e"}, {'Э', "E"}, {'ю', "yu"}, {'Ю', "Yu"},
                {'я', "ya"}, {'Я', "Ya"}
            }; // Словарь сопоставлений русских и английских букв
            char[] inChars = inputText.ToCharArray();
            string result = "";
            foreach (char c in inChars)
            {
                if (translitDictionary.ContainsKey(c))
                {
                    result += translitDictionary[c];
                }
                else
                {
                    result += c;
                }
            }
            return result;
        }
        // Проверка строки в массиве
        private bool containsInArr(string[] arr, string str, ref int index)
        {
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i].ToUpper() == str.ToUpper())
                {
                    index = i;
                    return true;
                }
            }
            return false;
        }
        // В поле выбранного пользователя нажата клавиша
        private void dnUserForChange_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btSelectUser.IsEnabled = false;
                btChangeSurname.IsEnabled = false;
                string searchStr = dnUserForChange.Text;
                Thread t = new Thread(() =>
                {
                    DirectorySearcher dirSearcher = new DirectorySearcher(_ADSession);
                    dirSearcher.SearchScope = SearchScope.Subtree;
                    dirSearcher.Filter = string.Format("(&(objectClass=user)(!(objectClass=computer))(|(cn=*{0}*)(sn=*{0}*)(givenName=*{0}*)(sAMAccountName=*{0}*)(distinguishedName={0})))", searchStr.Trim());
                    dirSearcher.PropertiesToLoad.Add("distinguishedName");
                    SearchResultCollection searchResults = dirSearcher.FindAll();
                    if (searchResults.Count > 0)
                    {
                        if (searchResults.Count > 1)
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                btSelectUser.IsEnabled = true;
                                btChangeSurname.IsEnabled = true;
                                btSelectUser_Click("dnUserForChange", new RoutedEventArgs());
                            }));
                        }
                        else
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                dnUserForChange.Text = (string)searchResults[0].Properties["distinguishedName"][0];
                                _userIsSelected = true;
                                dnUserForChange.FontWeight = FontWeights.Bold;
                                dnUserForChange.TextDecorations = TextDecorations.Underline;
                                btSelectUser.IsEnabled = true;
                                btChangeSurname.IsEnabled = true;
                            }));
                        }
                    }
                    else
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            btSelectUser.IsEnabled = true;
                            btChangeSurname.IsEnabled = true;
                            btSelectUser_Click("dnUserForChange", new RoutedEventArgs());
                        }));
                    }
                });
                t.IsBackground = true;
                t.Start();
            }
        }
        // Изменено значение поля выбранного пользователя
        private void dnUserForChange_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_userIsSelected)
            {
                _userIsSelected = false;
                dnUserForChange.FontWeight = FontWeights.Normal;
                dnUserForChange.TextDecorations = null;
            }
        }
    }
}
